from __future__ import annotations

import argparse
import json
import math
import re
import time
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence, Tuple
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen

EARTH_RADIUS_METERS = 6_371_000.0

# Coordinate ordering follows GeoJSON/OSRM: (lon, lat)
Coordinate = Tuple[float, float]


@dataclass
class RouteResult:
    mode: str
    code: str
    distance_m: float
    duration_s: float
    geometry: List[Coordinate]
    segments: List[Dict[str, float]] = field(default_factory=list)
    request_url: Optional[str] = None
    metadata: Dict[str, Any] = field(default_factory=dict)
    raw_response: Optional[Dict[str, Any]] = None

    def to_dict(self) -> Dict[str, Any]:
        return {
            "mode": self.mode,
            "code": self.code,
            "distance": self.distance_m,
            "duration": self.duration_s,
            "geometry": {
                "type": "LineString",
                "coordinates": self.geometry,
            },
            "segments": self.segments,
            "request_url": self.request_url,
            "metadata": self.metadata,
            "raw_response": self.raw_response,
        }


SHIPPING_WAYPOINTS: List[Dict[str, float | str]] = [
    {"id": "singapore", "lat": 1.29, "lon": 103.85},
    {"id": "strait_malacca", "lat": 3.0, "lon": 100.5},
    {"id": "andaman_sea", "lat": 9.0, "lon": 94.0},
    {"id": "taiwan_strait", "lat": 23.5, "lon": 119.8},
    {"id": "luzon_strait", "lat": 20.8, "lon": 122.3},
    {"id": "philippine_sea", "lat": 18.0, "lon": 132.0},
    {"id": "okinawa_approach", "lat": 26.0, "lon": 129.5},
    {"id": "bay_of_bengal", "lat": 13.0, "lon": 87.0},
    {"id": "arabian_sea", "lat": 15.0, "lon": 63.0},
    {"id": "gulf_aden", "lat": 12.0, "lon": 46.0},
    {"id": "red_sea_south", "lat": 17.0, "lon": 41.0},
    {"id": "suez_south", "lat": 29.4, "lon": 32.5},
    {"id": "suez_north", "lat": 31.2, "lon": 32.3},
    {"id": "med_east", "lat": 34.0, "lon": 26.0},
    {"id": "med_west", "lat": 37.0, "lon": 10.0},
    {"id": "gibraltar", "lat": 36.0, "lon": -5.5},
    {"id": "north_atlantic_east", "lat": 36.0, "lon": -20.0},
    {"id": "north_atlantic_mid", "lat": 38.0, "lon": -40.0},
    {"id": "north_atlantic_west", "lat": 36.0, "lon": -60.0},
    {"id": "us_east", "lat": 35.0, "lon": -74.0},
    {"id": "boston_approach", "lat": 42.6, "lon": -69.8},
    {"id": "caribbean", "lat": 17.0, "lon": -75.0},
    {"id": "panama_atlantic", "lat": 9.4, "lon": -79.8},
    {"id": "panama_pacific", "lat": 8.8, "lon": -79.6},
    {"id": "mexico_pacific", "lat": 18.0, "lon": -108.0},
    {"id": "california_approach", "lat": 33.0, "lon": -122.0},
    {"id": "south_china_sea", "lat": 14.0, "lon": 114.0},
    {"id": "east_china_sea", "lat": 28.0, "lon": 126.0},
    {"id": "japan_pacific", "lat": 35.0, "lon": 143.0},
    {"id": "north_pacific_west", "lat": 39.0, "lon": 165.0},
    {"id": "north_pacific_mid", "lat": 40.0, "lon": -170.0},
    {"id": "north_pacific_east", "lat": 37.0, "lon": -145.0},
    {"id": "aleutian_west", "lat": 50.0, "lon": 175.0},
    {"id": "aleutian_mid", "lat": 52.0, "lon": -175.0},
    {"id": "aleutian_east", "lat": 52.0, "lon": -160.0},
    {"id": "hawaii_north", "lat": 28.0, "lon": -158.0},
    {"id": "hawaii_east", "lat": 24.0, "lon": -145.0},
    {"id": "us_west_north", "lat": 47.0, "lon": -125.0},
    {"id": "us_west_mid", "lat": 38.0, "lon": -124.0},
    {"id": "baja_pacific", "lat": 25.0, "lon": -116.0},
    {"id": "equatorial_west_pacific", "lat": 4.0, "lon": 145.0},
    {"id": "equatorial_mid_pacific", "lat": 2.0, "lon": -170.0},
    {"id": "equatorial_east_pacific", "lat": 5.0, "lon": -125.0},
    {"id": "indian_ocean_mid", "lat": -20.0, "lon": 72.0},
    {"id": "cape_good_hope", "lat": -35.0, "lon": 18.0},
    {"id": "south_atlantic_mid", "lat": -20.0, "lon": -20.0},
]

SHIPPING_EDGES: List[Tuple[str, str]] = [
    ("singapore", "strait_malacca"),
    ("singapore", "south_china_sea"),
    ("south_china_sea", "taiwan_strait"),
    ("taiwan_strait", "east_china_sea"),
    ("taiwan_strait", "luzon_strait"),
    ("luzon_strait", "philippine_sea"),
    ("east_china_sea", "okinawa_approach"),
    ("okinawa_approach", "japan_pacific"),
    ("philippine_sea", "japan_pacific"),
    ("philippine_sea", "north_pacific_west"),
    ("equatorial_west_pacific", "philippine_sea"),
    ("strait_malacca", "andaman_sea"),
    ("andaman_sea", "bay_of_bengal"),
    ("bay_of_bengal", "arabian_sea"),
    ("arabian_sea", "gulf_aden"),
    ("gulf_aden", "red_sea_south"),
    ("red_sea_south", "suez_south"),
    ("suez_south", "suez_north"),
    ("suez_north", "med_east"),
    ("med_east", "med_west"),
    ("med_west", "gibraltar"),
    ("gibraltar", "north_atlantic_east"),
    ("north_atlantic_east", "north_atlantic_mid"),
    ("north_atlantic_mid", "north_atlantic_west"),
    ("north_atlantic_west", "us_east"),
    ("north_atlantic_west", "boston_approach"),
    ("us_east", "boston_approach"),
    ("north_atlantic_west", "caribbean"),
    ("caribbean", "panama_atlantic"),
    ("panama_atlantic", "panama_pacific"),
    ("panama_pacific", "mexico_pacific"),
    ("mexico_pacific", "california_approach"),
    ("south_china_sea", "east_china_sea"),
    ("east_china_sea", "japan_pacific"),
    ("japan_pacific", "north_pacific_west"),
    ("japan_pacific", "aleutian_west"),
    ("aleutian_west", "aleutian_mid"),
    ("aleutian_mid", "aleutian_east"),
    ("aleutian_east", "us_west_north"),
    ("us_west_north", "us_west_mid"),
    ("us_west_mid", "california_approach"),
    ("north_pacific_west", "north_pacific_mid"),
    ("north_pacific_mid", "north_pacific_east"),
    ("north_pacific_east", "california_approach"),
    ("north_pacific_west", "hawaii_north"),
    ("hawaii_north", "hawaii_east"),
    ("hawaii_east", "north_pacific_east"),
    ("hawaii_east", "equatorial_east_pacific"),
    ("south_china_sea", "equatorial_west_pacific"),
    ("equatorial_west_pacific", "equatorial_mid_pacific"),
    ("equatorial_mid_pacific", "equatorial_east_pacific"),
    ("equatorial_east_pacific", "panama_pacific"),
    ("arabian_sea", "indian_ocean_mid"),
    ("indian_ocean_mid", "cape_good_hope"),
    ("cape_good_hope", "south_atlantic_mid"),
    ("south_atlantic_mid", "north_atlantic_west"),
    ("south_atlantic_mid", "panama_atlantic"),
    ("mexico_pacific", "baja_pacific"),
    ("baja_pacific", "california_approach"),
]

AIR_QUERY_HINTS = (
    "airport",
    "international airport",
    "airfield",
    "機場",
    "國際機場",
)

SEA_QUERY_HINTS = (
    "port",
    "seaport",
    "harbor",
    "harbour",
    "terminal",
    "港",
    "港口",
    "碼頭",
)


def _to_rad(degrees: float) -> float:
    return (degrees * math.pi) / 180.0


def _clamp(value: float, min_value: float, max_value: float) -> float:
    return max(min_value, min(max_value, value))


def _normalize_text(value: str) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[\s\-_/,;:]+", " ", text)
    text = re.sub(r"[^\w\u4e00-\u9fff ]+", "", text)
    return re.sub(r"\s+", " ", text).strip()


def transport_type_to_mode(transport_type: str) -> str:
    normalized = _normalize_text(transport_type)
    if normalized in {"road", "rord", "road transport", "local land transport", "land", "truck", "express"}:
        return "driving"
    if normalized in {"air", "air transport"}:
        return "aviation-gc"
    if normalized in {"sea", "sea transport", "shipping", "marine", "ocean"}:
        return "shipping-gc"
    raise ValueError(f"Unsupported transport type: {transport_type}")


def _transport_type_to_hint(transport_type: str) -> Optional[str]:
    mode = transport_type_to_mode(transport_type)
    if mode == "aviation-gc":
        return "air"
    if mode == "shipping-gc":
        return "sea"
    return None


def _score_candidate(query: str, label: str, hint: Optional[str]) -> float:
    query_norm = _normalize_text(query)
    label_norm = _normalize_text(label)
    if not query_norm or not label_norm:
        return 0.0
    if query_norm == label_norm:
        score = 100.0
    elif query_norm in label_norm or label_norm in query_norm:
        score = 92.0
    else:
        query_tokens = set(query_norm.split())
        label_tokens = set(label_norm.split())
        overlap = len(query_tokens & label_tokens)
        score = (overlap / max(1, len(query_tokens))) * 75.0
        if all(token in label_norm for token in query_tokens):
            score = max(score, 88.0)
    hint_keywords = AIR_QUERY_HINTS if hint == "air" else SEA_QUERY_HINTS if hint == "sea" else ()
    if hint_keywords and any(keyword in label_norm for keyword in hint_keywords):
        score += 8.0
    return score


def _request_json(url: str, params: Dict[str, Any], timeout_sec: float) -> Any:
    request_url = f"{url}?{urlencode(params)}"
    request = Request(
        request_url,
        headers={
            "Accept": "application/json",
            "User-Agent": "transport-distance/1.0",
        },
    )
    with urlopen(request, timeout=timeout_sec) as response:
        payload = response.read().decode("utf-8", errors="replace")
    return json.loads(payload)


def _search_nominatim(query: str, timeout_sec: float, limit: int = 5) -> List[Dict[str, Any]]:
    data = _request_json(
        "https://nominatim.openstreetmap.org/search",
        {
            "q": query,
            "format": "jsonv2",
            "limit": limit,
            "addressdetails": 1,
        },
        timeout_sec=timeout_sec,
    )
    if not isinstance(data, list):
        return []
    candidates: List[Dict[str, Any]] = []
    for item in data:
        if not isinstance(item, dict):
            continue
        try:
            lat = float(item["lat"])
            lon = float(item["lon"])
        except (KeyError, TypeError, ValueError):
            continue
        label = str(item.get("display_name") or item.get("name") or query)
        candidates.append(
            {
                "lat": lat,
                "lon": lon,
                "label": label,
                "provider": "nominatim",
            }
        )
    return candidates


def _search_photon(query: str, timeout_sec: float, limit: int = 5) -> List[Dict[str, Any]]:
    data = _request_json(
        "https://photon.komoot.io/api",
        {
            "q": query,
            "limit": limit,
        },
        timeout_sec=timeout_sec,
    )
    if not isinstance(data, dict):
        return []
    features = data.get("features")
    if not isinstance(features, list):
        return []
    candidates: List[Dict[str, Any]] = []
    for feature in features:
        if not isinstance(feature, dict):
            continue
        geometry = feature.get("geometry")
        properties = feature.get("properties")
        if not isinstance(geometry, dict) or not isinstance(properties, dict):
            continue
        coords = geometry.get("coordinates")
        if not isinstance(coords, Sequence) or len(coords) < 2:
            continue
        try:
            lon = float(coords[0])
            lat = float(coords[1])
        except (TypeError, ValueError):
            continue
        label_parts = [
            properties.get("name"),
            properties.get("street"),
            properties.get("city"),
            properties.get("state"),
            properties.get("country"),
        ]
        label = ", ".join(str(part) for part in label_parts if part)
        candidates.append(
            {
                "lat": lat,
                "lon": lon,
                "label": label or query,
                "provider": "photon",
            }
        )
    return candidates


def _search_arcgis(query: str, timeout_sec: float, limit: int = 5) -> List[Dict[str, Any]]:
    data = _request_json(
        "https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/findAddressCandidates",
        {
            "f": "json",
            "SingleLine": query,
            "maxLocations": limit,
            "outFields": "Addr_type,Match_addr,Score",
        },
        timeout_sec=timeout_sec,
    )
    if not isinstance(data, dict):
        return []
    raw_candidates = data.get("candidates")
    if not isinstance(raw_candidates, list):
        return []

    candidates: List[Dict[str, Any]] = []
    for item in raw_candidates:
        if not isinstance(item, dict):
            continue
        location = item.get("location")
        if not isinstance(location, dict):
            continue
        try:
            lat = float(location["y"])
            lon = float(location["x"])
        except (KeyError, TypeError, ValueError):
            continue
        try:
            provider_score = float(item.get("score") or 0.0)
        except (TypeError, ValueError):
            provider_score = 0.0
        if provider_score < 65.0:
            continue
        attributes = item.get("attributes") if isinstance(item.get("attributes"), dict) else {}
        label = str(item.get("address") or attributes.get("Match_addr") or query)
        addr_type = str(attributes.get("Addr_type") or "")
        adjusted_score = provider_score
        if addr_type.lower() in {"locality", "postal", "postalext", "country", "region", "subregion", "district", "city", "neighborhood"}:
            adjusted_score -= 25.0
        elif addr_type.lower() in {"streetname", "streetint"}:
            adjusted_score += 10.0
        elif addr_type.lower() in {"pointaddress", "subaddress", "streetaddress", "poi"}:
            adjusted_score += 5.0
        candidates.append(
            {
                "lat": lat,
                "lon": lon,
                "label": label,
                "provider": "arcgis",
                "provider_score": adjusted_score,
                "raw_provider_score": provider_score,
                "addr_type": addr_type,
            }
        )
    return candidates


def _normalize_address_for_geocode(address: str) -> str:
    text = str(address or "").strip()
    if not text:
        return ""
    replacements = {
        "\u3000": " ",
        "\uff0c": ",",
        "\uff08": "(",
        "\uff09": ")",
        "\u2160": "I",
        "\u2161": "II",
        "\u2162": "III",
        "\u2163": "IV",
        "\u2164": "V",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = re.sub(r"[‐‑‒–—―]", "-", text)
    text = re.sub(r"(?<=[a-z])(?=[A-Z0-9])", " ", text)
    text = re.sub(r"(?<=\d)(?=[A-Za-z])", " ", text)
    text = re.sub(r"\b(road|rd|avenue|ave)\s+iii\b", r"\1 3", text, flags=re.IGNORECASE)
    text = re.sub(r"\b(road|rd|avenue|ave)\s+ii\b", r"\1 2", text, flags=re.IGNORECASE)
    text = re.sub(r"\s+,", ",", text)
    text = re.sub(r",+", ",", text)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip(" ,")


def _strip_address_unit_details(address: str) -> str:
    text = str(address or "")
    text = re.sub(r"\([^)]*\)", " ", text)
    text = re.sub(r"\b\d{1,2}(?:st|nd|rd|th)\s+floor\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\b(?:suite|ste|apt|apartment|unit|room|rm|fl)\b\.?\s*#?\s*[-\w/]+", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"#\s*[-\w/]+", " ", text)
    text = re.sub(r"\s+,", ",", text)
    text = re.sub(r",+", ",", text)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip(" ,")


def _expand_common_address_abbreviations(address: str) -> str:
    text = str(address or "")
    replacements = {
        r"\bind\.?\b": "industrial",
        r"\bave\.?\b": "avenue",
        r"\brd\.?\b": "road",
        r"\bblvd\.?\b": "boulevard",
        r"\bstn\.?\b": "station",
    }
    for pattern, replacement in replacements.items():
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
    return re.sub(r"\s{2,}", " ", text).strip()


def _cleanup_address_part(part: str) -> str:
    text = re.sub(r"\b\d{3,6}(?:-\d{3,4})?\b", " ", str(part or ""))
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip(" ,")


def _remove_leading_house_number(street: str) -> str:
    text = re.sub(r"^\s*no\.?\s*\d+[a-z-]*\s*", "", str(street or ""), flags=re.IGNORECASE)
    text = re.sub(r"^\s*\d+[a-z-]*\s+", "", text, flags=re.IGNORECASE)
    return re.sub(r"\s{2,}", " ", text).strip()


def _build_street_city_country_candidates(address: str) -> List[str]:
    parts = [_cleanup_address_part(part) for part in str(address or "").split(",")]
    parts = [part for part in parts if part]
    if len(parts) < 2:
        return []
    country = parts[-1]
    city = parts[-2] if len(parts) >= 3 else ""
    street = parts[0]
    candidates = []
    if street and city and country:
        candidates.append(f"{street}, {city}, {country}")
        no_house_number = _remove_leading_house_number(street)
        if no_house_number and no_house_number.lower() != street.lower():
            candidates.append(f"{no_house_number}, {city}, {country}")
    if city and country:
        candidates.append(f"{city}, {country}")
    return [_normalize_address_for_geocode(item) for item in candidates]


def _looks_like_detailed_address(query: str) -> bool:
    text = str(query or "")
    if not text.strip():
        return False
    if re.search(r"\b(c/o|warehouse|dock|suite|road|rd|street|st|avenue|ave|drive|dr|floor|building|bldg|lot)\b", text, re.IGNORECASE):
        return True
    if re.search(r"\d{1,6}\s*[A-Za-z]", text) and "," in text:
        return True
    return False


def _known_address_fallback_queries(query: str) -> List[str]:
    normalized = _normalize_address_for_geocode(query).lower()
    fallbacks = []
    if "atl logistics" in normalized and "kwai chung" in normalized:
        fallbacks.append("ATL Logistics Centre A, Berth 3, Kwai Chung Container Terminal, Kwai Chung, Hong Kong")
        fallbacks.append("Kwai Chung Container Terminal, Kwai Chung, Hong Kong")
    if "ping ha road" in normalized and "yuen long" in normalized:
        fallbacks.append("Ping Ha Road, Lau Fau Shan, Yuen Long, Hong Kong")
    if "marathahalli" in normalized and ("outer ring road" in normalized or "salarpuria" in normalized):
        fallbacks.append("Outer Ring Road, Marathahalli, Bangalore, India")
        fallbacks.append("Salarpuria Supreme, Marathahalli, Bangalore, India")
    return fallbacks


def _build_query_variants(query: str, hint: Optional[str]) -> List[str]:
    base = str(query or "").strip()
    if not base:
        return []
    normalized = _normalize_text(base)
    normalized_address = _normalize_address_for_geocode(base)
    relaxed_address = _strip_address_unit_details(normalized_address)
    expanded_address = _expand_common_address_abbreviations(relaxed_address or normalized_address)
    variants = [
        item
        for item in (
            relaxed_address,
            expanded_address,
            *_known_address_fallback_queries(base),
            *_build_street_city_country_candidates(expanded_address or relaxed_address or normalized_address),
            normalized_address,
            base,
        )
        if item
    ]
    hint_keywords = AIR_QUERY_HINTS if hint == "air" else SEA_QUERY_HINTS if hint == "sea" else ()
    if hint_keywords and not any(keyword in normalized for keyword in hint_keywords):
        if hint == "air":
            variants.extend([f"{base} airport", f"{base} international airport"])
        elif hint == "sea":
            variants.extend([f"{base} port", f"Port of {base}"])
    deduped: List[str] = []
    seen = set()
    for item in variants:
        key = item.strip().lower()
        if key and key not in seen:
            seen.add(key)
            deduped.append(item)
        if len(deduped) >= 20:
            break
    return deduped


def geocode_place(
    query: str,
    *,
    transport_type: str = "",
    timeout_sec: float = 10.0,
    cache: Optional[Dict[Tuple[str, str], Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    hint = _transport_type_to_hint(transport_type) if transport_type else None
    cache_key = (hint or "", _normalize_text(query))
    if cache is not None and cache_key in cache:
        return cache[cache_key]

    variants = _build_query_variants(query, hint)
    candidates: List[Dict[str, Any]] = []
    for variant in variants:
        for search_fn in (_search_nominatim, _search_photon):
            try:
                found_items = search_fn(variant, timeout_sec=timeout_sec)
            except Exception:
                continue
            for item in found_items:
                item["score"] = _score_candidate(variant, str(item.get("label") or ""), hint)
                item["query_used"] = variant
                candidates.append(item)

    best_primary_score = max((float(item.get("score") or 0.0) for item in candidates), default=0.0)
    if not candidates or best_primary_score < 85.0 or _looks_like_detailed_address(query):
        for variant in variants:
            try:
                found_items = _search_arcgis(variant, timeout_sec=timeout_sec)
            except Exception:
                continue
            for item in found_items:
                item["score"] = float(item.get("provider_score") or 0.0)
                item["query_used"] = variant
                candidates.append(item)

    if not candidates:
        raise RuntimeError(f"Could not resolve location: {query}")

    best = max(candidates, key=lambda item: float(item.get("score") or 0.0))
    resolved = {
        "query": query,
        "query_used": str(best.get("query_used") or query),
        "lat": float(best["lat"]),
        "lon": float(best["lon"]),
        "label": str(best.get("label") or query),
        "provider": str(best.get("provider") or ""),
    }
    if cache is not None:
        cache[cache_key] = resolved
    return resolved


def haversine_meters(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    d_lat = _to_rad(lat2 - lat1)
    d_lon = _to_rad(lon2 - lon1)
    a = (
        math.sin(d_lat / 2) * math.sin(d_lat / 2)
        + math.cos(_to_rad(lat1)) * math.cos(_to_rad(lat2)) * math.sin(d_lon / 2) * math.sin(d_lon / 2)
    )
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return EARTH_RADIUS_METERS * c


def _to_cartesian(lat: float, lon: float) -> Tuple[float, float, float]:
    lat_rad = _to_rad(lat)
    lon_rad = _to_rad(lon)
    cos_lat = math.cos(lat_rad)
    return (cos_lat * math.cos(lon_rad), cos_lat * math.sin(lon_rad), math.sin(lat_rad))


def _from_cartesian(point: Tuple[float, float, float]) -> Tuple[float, float]:
    x, y, z = point
    hyp = math.sqrt((x * x) + (y * y))
    lat = math.atan2(z, hyp)
    lon = math.atan2(y, x)
    return ((lat * 180.0) / math.pi, (lon * 180.0) / math.pi)


def _lerp_great_circle_point(from_lat: float, from_lon: float, to_lat: float, to_lon: float, t: float) -> Tuple[float, float]:
    sx, sy, sz = _to_cartesian(from_lat, from_lon)
    ex, ey, ez = _to_cartesian(to_lat, to_lon)
    dot = _clamp((sx * ex) + (sy * ey) + (sz * ez), -1.0, 1.0)
    omega = math.acos(dot)
    if omega < 1e-12:
        return (from_lat, from_lon)

    sin_omega = math.sin(omega)
    scale_start = math.sin((1 - t) * omega) / sin_omega
    scale_end = math.sin(t * omega) / sin_omega
    px = (scale_start * sx) + (scale_end * ex)
    py = (scale_start * sy) + (scale_end * ey)
    pz = (scale_start * sz) + (scale_end * ez)
    return _from_cartesian((px, py, pz))


def build_aviation_great_circle_route(
    from_lat: float,
    from_lon: float,
    to_lat: float,
    to_lon: float,
    speed_kph: float = 900.0,
    segment_max_km: Optional[float] = 500.0,
) -> RouteResult:
    if speed_kph <= 0:
        raise ValueError("speed_kph must be positive")

    distance = haversine_meters(from_lat, from_lon, to_lat, to_lon)
    speed_mps = (speed_kph * 1000.0) / 3600.0
    coordinates: List[Coordinate] = []
    segments: List[Dict[str, float]] = []

    if distance == 0 or not segment_max_km:
        coordinates = [(from_lon, from_lat), (to_lon, to_lat)]
    else:
        capped_segment_km = _clamp(float(segment_max_km), 1.0, 5000.0)
        segment_count = max(2, math.ceil((distance / 1000.0) / capped_segment_km))
        for i in range(segment_count + 1):
            t = i / segment_count
            lat, lon = _lerp_great_circle_point(from_lat, from_lon, to_lat, to_lon, t)
            coordinates.append((lon, lat))

        for idx in range(len(coordinates) - 1):
            lon1, lat1 = coordinates[idx]
            lon2, lat2 = coordinates[idx + 1]
            segment_distance = haversine_meters(lat1, lon1, lat2, lon2)
            segments.append(
                {
                    "segment": float(idx + 1),
                    "distance": segment_distance,
                    "duration": segment_distance / speed_mps,
                }
            )

    return RouteResult(
        mode="aviation-gc",
        code="AviationGreatCircle",
        distance_m=distance,
        duration_s=distance / speed_mps,
        geometry=coordinates,
        segments=segments,
        metadata={"speed_kph": speed_kph, "segment_max_km": segment_max_km},
    )


def _build_adjacency(nodes: Dict[str, Dict[str, float | str]]) -> Dict[str, List[Tuple[str, float]]]:
    adjacency: Dict[str, List[Tuple[str, float]]] = {node_id: [] for node_id in nodes}
    for a, b in SHIPPING_EDGES:
        node_a = nodes.get(a)
        node_b = nodes.get(b)
        if not node_a or not node_b:
            continue
        dist = haversine_meters(
            float(node_a["lat"]),
            float(node_a["lon"]),
            float(node_b["lat"]),
            float(node_b["lon"]),
        )
        adjacency[a].append((b, dist))
        adjacency[b].append((a, dist))
    return adjacency


def _shortest_path_dijkstra(adjacency: Dict[str, List[Tuple[str, float]]], start_id: str, end_id: str) -> Optional[List[str]]:
    dist: Dict[str, float] = {node_id: math.inf for node_id in adjacency}
    prev: Dict[str, str] = {}
    visited: set[str] = set()
    dist[start_id] = 0.0

    while len(visited) < len(adjacency):
        current = None
        current_dist = math.inf
        for node_id, node_dist in dist.items():
            if node_id not in visited and node_dist < current_dist:
                current = node_id
                current_dist = node_dist

        if current is None or current == end_id:
            break

        visited.add(current)
        for neighbor, edge_cost in adjacency.get(current, []):
            if neighbor in visited:
                continue
            alt = current_dist + edge_cost
            if alt < dist[neighbor]:
                dist[neighbor] = alt
                prev[neighbor] = current

    if not math.isfinite(dist.get(end_id, math.inf)):
        return None

    path = [end_id]
    while path[-1] != start_id:
        parent = prev.get(path[-1])
        if parent is None:
            return None
        path.append(parent)
    path.reverse()
    return path


def _connect_virtual_node(
    adjacency: Dict[str, List[Tuple[str, float]]],
    nodes: Dict[str, Dict[str, float | str]],
    virtual_id: str,
    max_links: int,
    max_distance_km: float,
) -> None:
    virtual = nodes[virtual_id]
    ranked: List[Tuple[str, float]] = []
    for wp in SHIPPING_WAYPOINTS:
        wp_id = str(wp["id"])
        d = haversine_meters(float(virtual["lat"]), float(virtual["lon"]), float(wp["lat"]), float(wp["lon"]))
        ranked.append((wp_id, d))
    ranked.sort(key=lambda x: x[1])

    threshold_m = _clamp(max_distance_km, 50.0, 20_000.0) * 1000.0
    in_range = [item for item in ranked if item[1] <= threshold_m]
    selected = (in_range if in_range else ranked[:1])[:max_links]

    for node_id, cost in selected:
        adjacency[virtual_id].append((node_id, cost))
        adjacency[node_id].append((virtual_id, cost))


def build_shipping_navigable_route(
    from_lat: float,
    from_lon: float,
    to_lat: float,
    to_lon: float,
    speed_kph: float = 35.0,
) -> RouteResult:
    if speed_kph <= 0:
        raise ValueError("speed_kph must be positive")

    fallback = build_aviation_great_circle_route(
        from_lat=from_lat,
        from_lon=from_lon,
        to_lat=to_lat,
        to_lon=to_lon,
        speed_kph=speed_kph,
        segment_max_km=None,
    )
    fallback.mode = "shipping-gc"
    fallback.code = "ShippingGreatCircleFallback"

    nodes: Dict[str, Dict[str, float | str]] = {
        str(wp["id"]): {"id": str(wp["id"]), "lat": float(wp["lat"]), "lon": float(wp["lon"])} for wp in SHIPPING_WAYPOINTS
    }
    nodes["_start"] = {"id": "_start", "lat": from_lat, "lon": from_lon}
    nodes["_end"] = {"id": "_end", "lat": to_lat, "lon": to_lon}

    adjacency = _build_adjacency(nodes)
    _connect_virtual_node(adjacency, nodes, "_start", max_links=4, max_distance_km=1800.0)
    _connect_virtual_node(adjacency, nodes, "_end", max_links=4, max_distance_km=1800.0)

    path = _shortest_path_dijkstra(adjacency, "_start", "_end")
    if not path or len(path) < 2:
        fallback.metadata = {"shipping_mode": "great-circle-fallback"}
        return fallback

    coordinates: List[Coordinate] = []
    for node_id in path:
        node = nodes[node_id]
        coordinates.append((float(node["lon"]), float(node["lat"])))

    distance = 0.0
    for i in range(1, len(coordinates)):
        lon1, lat1 = coordinates[i - 1]
        lon2, lat2 = coordinates[i]
        distance += haversine_meters(lat1, lon1, lat2, lon2)

    speed_mps = (speed_kph * 1000.0) / 3600.0
    return RouteResult(
        mode="shipping-gc",
        code="ShippingNavigableApprox",
        distance_m=distance,
        duration_s=distance / speed_mps,
        geometry=coordinates,
        segments=[],
        metadata={
            "shipping_mode": "navigable-approx",
            "shipping_node_path": path,
            "speed_kph": speed_kph,
        },
    )


def build_osrm_request_url(
    api_base: str,
    profile: str,
    from_lat: float,
    from_lon: float,
    to_lat: float,
    to_lon: float,
    alternatives: bool = False,
    steps: bool = False,
    annotations: bool = False,
) -> str:
    base = api_base.strip().rstrip("/")
    if not base:
        raise ValueError("api_base is required")
    if not profile:
        raise ValueError("profile is required")

    params = urlencode(
        {
            "overview": "full",
            "geometries": "geojson",
            "alternatives": str(bool(alternatives)).lower(),
            "steps": str(bool(steps)).lower(),
            "annotations": str(bool(annotations)).lower(),
        }
    )
    path = f"/route/v1/{quote(profile, safe='')}/{from_lon},{from_lat};{to_lon},{to_lat}"
    return f"{base}{path}?{params}"


def fetch_osrm_route(
    from_lat: float,
    from_lon: float,
    to_lat: float,
    to_lon: float,
    api_base: str = "https://router.project-osrm.org",
    profile: str = "driving",
    alternatives: bool = False,
    steps: bool = False,
    annotations: bool = False,
    timeout_sec: float = 20.0,
    retry_count: int = 3,
    retry_delay_sec: float = 0.8,
) -> RouteResult:
    request_url = build_osrm_request_url(
        api_base=api_base,
        profile=profile,
        from_lat=from_lat,
        from_lon=from_lon,
        to_lat=to_lat,
        to_lon=to_lon,
        alternatives=alternatives,
        steps=steps,
        annotations=annotations,
    )
    req = Request(
        request_url,
        headers={
            "Accept": "application/json",
            "User-Agent": "transport-distance/1.0",
        },
    )
    last_error = None
    for attempt in range(max(1, retry_count)):
        try:
            with urlopen(req, timeout=timeout_sec) as resp:
                status_code = getattr(resp, "status", 200)
                payload = resp.read().decode("utf-8", errors="replace")
            break
        except Exception as exc:
            last_error = exc
            if attempt + 1 >= max(1, retry_count):
                raise RuntimeError(f"OSRM request failed: {exc}") from exc
            time.sleep(retry_delay_sec)
    else:
        raise RuntimeError(f"OSRM request failed: {last_error}")

    try:
        data = json.loads(payload)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"OSRM returned non-JSON response (HTTP {status_code})") from exc

    if status_code < 200 or status_code >= 300:
        msg = data.get("message") if isinstance(data, dict) else None
        raise RuntimeError(f"OSRM request failed: {msg or f'HTTP {status_code}'}")

    if not isinstance(data, dict) or data.get("code") != "Ok":
        msg = data.get("message") if isinstance(data, dict) else None
        raise RuntimeError(f"OSRM request failed: {msg or 'Unexpected response code'}")

    routes = data.get("routes")
    if not isinstance(routes, Sequence) or not routes:
        raise RuntimeError("OSRM response has no routes")

    route = routes[0]
    if not isinstance(route, dict):
        raise RuntimeError("OSRM route payload is invalid")

    geometry = route.get("geometry", {})
    coordinates = geometry.get("coordinates") if isinstance(geometry, dict) else None
    if not isinstance(coordinates, Sequence) or len(coordinates) < 2:
        raise RuntimeError("OSRM route geometry is missing or invalid")

    try:
        distance = float(route["distance"])
        duration = float(route["duration"])
    except (KeyError, TypeError, ValueError) as exc:
        raise RuntimeError("OSRM route distance/duration missing") from exc

    normalized_coords: List[Coordinate] = []
    for item in coordinates:
        if not isinstance(item, Sequence) or len(item) < 2:
            continue
        normalized_coords.append((float(item[0]), float(item[1])))

    if len(normalized_coords) < 2:
        raise RuntimeError("OSRM route geometry coordinates are invalid")

    return RouteResult(
        mode="osrm",
        code=str(data.get("code")),
        distance_m=distance,
        duration_s=duration,
        geometry=normalized_coords,
        segments=[],
        request_url=request_url,
        metadata={
            "profile": profile,
            "alternatives": bool(alternatives),
            "steps": bool(steps),
            "annotations": bool(annotations),
            "api_base": api_base,
        },
        raw_response=data,
    )


def compute_transport_distance(
    mode: str,
    from_lat: float,
    from_lon: float,
    to_lat: float,
    to_lon: float,
    *,
    api_base: str = "https://router.project-osrm.org",
    profile: str = "driving",
    alternatives: bool = False,
    steps: bool = False,
    annotations: bool = False,
    aviation_speed_kph: float = 900.0,
    aviation_segment_max_km: Optional[float] = 500.0,
    shipping_speed_kph: float = 35.0,
    timeout_sec: float = 20.0,
) -> RouteResult:
    normalized_mode = (mode or "").strip().lower()

    if normalized_mode in {"aviation", "aviation-gc", "great-circle", "gc"}:
        return build_aviation_great_circle_route(
            from_lat=from_lat,
            from_lon=from_lon,
            to_lat=to_lat,
            to_lon=to_lon,
            speed_kph=aviation_speed_kph,
            segment_max_km=aviation_segment_max_km,
        )

    if normalized_mode in {"shipping", "shipping-gc", "shipping-navigable"}:
        return build_shipping_navigable_route(
            from_lat=from_lat,
            from_lon=from_lon,
            to_lat=to_lat,
            to_lon=to_lon,
            speed_kph=shipping_speed_kph,
        )

    # Accept either mode="osrm" with profile=..., or mode directly as an OSRM profile.
    osrm_profiles = {"driving", "driving-traffic", "walking", "cycling"}
    if normalized_mode in osrm_profiles:
        profile = normalized_mode
        normalized_mode = "osrm"

    if normalized_mode == "osrm":
        return fetch_osrm_route(
            from_lat=from_lat,
            from_lon=from_lon,
            to_lat=to_lat,
            to_lon=to_lon,
            api_base=api_base,
            profile=profile,
            alternatives=alternatives,
            steps=steps,
            annotations=annotations,
            timeout_sec=timeout_sec,
        )

    raise ValueError(
        "Unsupported mode. Use one of: osrm, driving, driving-traffic, walking, cycling, aviation-gc, shipping-gc."
    )


def compute_transport_distance_from_queries(
    transport_type: str,
    from_query: str,
    to_query: str,
    *,
    api_base: str = "https://router.project-osrm.org",
    timeout_sec: float = 20.0,
    geocode_timeout_sec: float = 10.0,
    geocode_cache: Optional[Dict[Tuple[str, str], Dict[str, Any]]] = None,
) -> RouteResult:
    mode = transport_type_to_mode(transport_type)
    from_place = geocode_place(
        from_query,
        transport_type=transport_type,
        timeout_sec=geocode_timeout_sec,
        cache=geocode_cache,
    )
    to_place = geocode_place(
        to_query,
        transport_type=transport_type,
        timeout_sec=geocode_timeout_sec,
        cache=geocode_cache,
    )
    result = compute_transport_distance(
        mode=mode,
        from_lat=float(from_place["lat"]),
        from_lon=float(from_place["lon"]),
        to_lat=float(to_place["lat"]),
        to_lon=float(to_place["lon"]),
        api_base=api_base,
        timeout_sec=timeout_sec,
    )
    result.metadata.update(
        {
            "transport_type": transport_type,
            "from_query": from_query,
            "to_query": to_query,
            "from_place": from_place,
            "to_place": to_place,
        }
    )
    return result


def mode_requires_network(mode: str, use_geocoding: bool = False) -> bool:
    normalized_mode = (mode or "").strip().lower()
    if normalized_mode in {"aviation", "aviation-gc", "great-circle", "gc", "shipping", "shipping-gc", "shipping-navigable"}:
        return bool(use_geocoding)
    return True


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Compute OSRM / aviation / shipping transport distance.")
    parser.add_argument("--mode", required=True, help="osrm | driving | walking | cycling | aviation-gc | shipping-gc")
    parser.add_argument("--from-lat", type=float, required=True)
    parser.add_argument("--from-lon", type=float, required=True)
    parser.add_argument("--to-lat", type=float, required=True)
    parser.add_argument("--to-lon", type=float, required=True)
    parser.add_argument("--api-base", default="https://router.project-osrm.org")
    parser.add_argument("--profile", default="driving")
    parser.add_argument("--alternatives", action="store_true")
    parser.add_argument("--steps", action="store_true")
    parser.add_argument("--annotations", action="store_true")
    parser.add_argument("--aviation-speed-kph", type=float, default=900.0)
    parser.add_argument("--aviation-segment-max-km", type=float, default=500.0)
    parser.add_argument("--shipping-speed-kph", type=float, default=35.0)
    parser.add_argument("--timeout-sec", type=float, default=20.0)
    return parser.parse_args()


def main() -> None:
    args = _parse_args()
    result = compute_transport_distance(
        mode=args.mode,
        from_lat=args.from_lat,
        from_lon=args.from_lon,
        to_lat=args.to_lat,
        to_lon=args.to_lon,
        api_base=args.api_base,
        profile=args.profile,
        alternatives=args.alternatives,
        steps=args.steps,
        annotations=args.annotations,
        aviation_speed_kph=args.aviation_speed_kph,
        aviation_segment_max_km=args.aviation_segment_max_km,
        shipping_speed_kph=args.shipping_speed_kph,
        timeout_sec=args.timeout_sec,
    )
    print(json.dumps(result.to_dict(), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
