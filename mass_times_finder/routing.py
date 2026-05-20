import requests
import math

NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
OSRM_URL = "https://router.project-osrm.org/route/v1/driving"
HEADERS = {"User-Agent": "MassTimesRouteFinder/1.0"}


def geocode(address):
    """Return (lat, lng) for address string, or None."""
    resp = requests.get(
        NOMINATIM_URL,
        params={"q": address, "format": "json", "limit": 1, "countrycodes": "us"},
        headers=HEADERS,
        timeout=10,
    )
    results = resp.json()
    if not results:
        return None
    return float(results[0]["lat"]), float(results[0]["lon"])


def get_route(start, end):
    """Return dict with coordinates list [[lat,lng],...], distance_miles, duration_hours."""
    lat1, lng1 = start
    lat2, lng2 = end
    url = f"{OSRM_URL}/{lng1},{lat1};{lng2},{lat2}"
    resp = requests.get(
        url,
        params={"overview": "full", "geometries": "geojson"},
        timeout=30,
    )
    data = resp.json()
    if data.get("code") != "Ok":
        return None
    route = data["routes"][0]
    coords = [[c[1], c[0]] for c in route["geometry"]["coordinates"]]  # [lng,lat] -> [lat,lng]
    return {
        "coordinates": coords,
        "distance_miles": round(route["distance"] / 1609.34, 1),
        "duration_hours": round(route["duration"] / 3600, 1),
    }


def haversine(lat1, lng1, lat2, lng2):
    R = 3959
    dlat = math.radians(lat2 - lat1)
    dlng = math.radians(lng2 - lng1)
    a = (
        math.sin(dlat / 2) ** 2
        + math.cos(math.radians(lat1))
        * math.cos(math.radians(lat2))
        * math.sin(dlng / 2) ** 2
    )
    return R * 2 * math.asin(math.sqrt(a))


def sample_route_points(coordinates, interval_miles=50):
    """Sample points from route polyline every interval_miles."""
    if not coordinates:
        return []
    points = [coordinates[0]]
    accumulated = 0.0
    for i in range(1, len(coordinates)):
        prev, curr = coordinates[i - 1], coordinates[i]
        accumulated += haversine(prev[0], prev[1], curr[0], curr[1])
        if accumulated >= interval_miles:
            points.append(curr)
            accumulated = 0.0
    if points[-1] != coordinates[-1]:
        points.append(coordinates[-1])
    return points
