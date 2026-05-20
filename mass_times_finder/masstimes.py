import requests
from bs4 import BeautifulSoup
import json

BROWSER_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}


def search_near_location(lat, lng, radius_miles=25, day="all"):
    """Return list of church dicts near (lat, lng)."""
    churches = []
    try:
        churches = _search_masstimes(lat, lng, radius_miles, day)
    except Exception as e:
        print(f"MassTimes.org error: {e}")
    if not churches:
        try:
            churches = _search_overpass(lat, lng, radius_miles)
        except Exception as e:
            print(f"Overpass error: {e}")
    return churches


def _search_masstimes(lat, lng, radius_miles, day):
    session = requests.Session()
    session.headers.update(BROWSER_HEADERS)

    # Try several URL patterns that MassTimes.org has used
    urls_to_try = [
        (
            "GET",
            "https://www.masstimes.org/search",
            {
                "lat": lat,
                "lng": lng,
                "distance": radius_miles,
                **({"day": day} if day != "all" else {}),
            },
        ),
        (
            "GET",
            "https://www.masstimes.org/api/v1/search",
            {"lat": lat, "lng": lng, "distance": radius_miles},
        ),
        (
            "GET",
            f"https://www.masstimes.org/mass-times/lat/{lat}/lng/{lng}",
            {},
        ),
    ]

    for method, url, params in urls_to_try:
        try:
            resp = session.get(url, params=params, timeout=12)
            if resp.status_code == 200:
                if "application/json" in resp.headers.get("Content-Type", ""):
                    return _parse_masstimes_json(resp.json(), day)
                else:
                    result = _parse_masstimes_html(resp.text, day)
                    if result:
                        return result
        except Exception:
            continue
    return []


def _parse_masstimes_html(html, day_filter="all"):
    soup = BeautifulSoup(html, "html.parser")
    churches = []

    # Try to extract embedded JSON data first
    for script in soup.find_all("script"):
        text = script.string or ""
        for pattern in [
            "var churches =",
            "var locations =",
            "window.data =",
            "var data =",
        ]:
            if pattern in text:
                try:
                    start = text.index(pattern) + len(pattern)
                    end = text.index(";", start)
                    data = json.loads(text[start:end].strip())
                    if isinstance(data, list):
                        return [
                            _normalize(c)
                            for c in data
                            if c.get("name") or c.get("parish_name")
                        ]
                except Exception:
                    pass

    # DOM scraping fallback
    selectors = [
        ".church-card",
        ".parish-card",
        ".result-card",
        ".location-result",
        "[class*='church']",
        "[class*='parish']",
        "li.result",
        ".search-result",
    ]
    for sel in selectors:
        items = soup.select(sel)
        if items:
            for item in items:
                c = _extract_church_from_element(item)
                if c:
                    churches.append(c)
            break

    return churches


def _extract_church_from_element(el):
    name_el = el.select_one("h1,h2,h3,h4,.name,.church-name,.parish-name")
    if not name_el:
        return None
    name = name_el.get_text(strip=True)
    if not name:
        return None

    addr_el = el.select_one(".address,.addr,[class*='address']")
    address = addr_el.get_text(strip=True) if addr_el else ""

    lat = el.get("data-lat") or el.get("data-latitude")
    lng = el.get("data-lng") or el.get("data-longitude") or el.get("data-lon")
    if not lat or not lng:
        return None

    times_els = el.select(".mass-time,.time,[class*='time'],[class*='mass']")
    mass_times = [t.get_text(strip=True) for t in times_els if t.get_text(strip=True)]

    return {
        "name": name,
        "lat": float(lat),
        "lng": float(lng),
        "address": address,
        "city": "",
        "state": "",
        "mass_times": mass_times,
        "phone": "",
        "website": "",
        "source": "masstimes.org",
    }


def _parse_masstimes_json(data, day_filter="all"):
    if isinstance(data, dict):
        data = data.get("churches", data.get("results", data.get("data", [])))
    return [_normalize(c) for c in (data or []) if isinstance(c, dict)]


def _normalize(data):
    return {
        "name": data.get(
            "name", data.get("parish_name", data.get("title", "Catholic Church"))
        ),
        "lat": float(data.get("lat", data.get("latitude", 0))),
        "lng": float(data.get("lng", data.get("longitude", data.get("lon", 0)))),
        "address": data.get("address", data.get("street", "")),
        "city": data.get("city", ""),
        "state": data.get("state", ""),
        "mass_times": data.get("mass_times", data.get("times", [])),
        "phone": data.get("phone", data.get("telephone", "")),
        "website": data.get("website", data.get("url", "")),
        "source": "masstimes.org",
    }


def _search_overpass(lat, lng, radius_miles=25):
    radius_m = int(radius_miles * 1609.34)
    query = f"""[out:json][timeout:30];
(
  node["amenity"="place_of_worship"]["religion"="christian"]["denomination"="catholic"](around:{radius_m},{lat},{lng});
  way["amenity"="place_of_worship"]["religion"="christian"]["denomination"="catholic"](around:{radius_m},{lat},{lng});
  node["amenity"="place_of_worship"]["denomination"="roman_catholic"](around:{radius_m},{lat},{lng});
  way["amenity"="place_of_worship"]["denomination"="roman_catholic"](around:{radius_m},{lat},{lng});
);
out center tags;"""

    resp = requests.post(
        "https://overpass-api.de/api/interpreter",
        data={"data": query},
        timeout=35,
    )
    resp.raise_for_status()
    data = resp.json()

    churches = []
    seen = set()
    for el in data.get("elements", []):
        tags = el.get("tags", {})
        name = tags.get("name", "")
        if not name or name in seen:
            continue
        seen.add(name)

        if el["type"] == "node":
            elat, elng = el["lat"], el["lon"]
        else:
            center = el.get("center", {})
            elat, elng = center.get("lat", 0), center.get("lon", 0)
        if not elat or not elng:
            continue

        house = tags.get("addr:housenumber", "")
        street = tags.get("addr:street", "")
        address = f"{house} {street}".strip()

        churches.append(
            {
                "name": name,
                "lat": elat,
                "lng": elng,
                "address": address,
                "city": tags.get("addr:city", ""),
                "state": tags.get("addr:state", ""),
                "mass_times": [],
                "phone": tags.get("phone", tags.get("contact:phone", "")),
                "website": tags.get("website", tags.get("contact:website", "")),
                "source": "openstreetmap",
            }
        )

    return churches
