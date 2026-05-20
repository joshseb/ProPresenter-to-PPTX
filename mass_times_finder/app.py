from flask import Flask, render_template, request, jsonify
from routing import geocode, get_route, sample_route_points
from masstimes import search_near_location

app = Flask(__name__)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/search", methods=["POST"])
def search():
    data = request.json or {}
    start_addr = data.get("start", "").strip()
    end_addr = data.get("end", "").strip()
    day = data.get("day", "all")
    radius = int(data.get("radius", 25))
    interval = int(data.get("interval", 50))

    if not start_addr or not end_addr:
        return jsonify({"error": "Start and end addresses are required"}), 400

    start_coords = geocode(start_addr)
    if not start_coords:
        return jsonify({"error": f"Could not find: {start_addr}"}), 400

    end_coords = geocode(end_addr)
    if not end_coords:
        return jsonify({"error": f"Could not find: {end_addr}"}), 400

    route = get_route(start_coords, end_coords)
    if not route:
        return jsonify({"error": "Could not calculate route"}), 500

    sample_points = sample_route_points(route["coordinates"], interval)

    all_churches = []
    seen = set()
    for lat, lng in sample_points:
        for church in search_near_location(lat, lng, radius, day):
            uid = f"{church['name']}|{round(church['lat'], 3)}|{round(church['lng'], 3)}"
            if uid not in seen:
                seen.add(uid)
                all_churches.append(church)

    return jsonify(
        {
            "route": route,
            "churches": all_churches,
            "start": {
                "address": start_addr,
                "lat": start_coords[0],
                "lng": start_coords[1],
            },
            "end": {
                "address": end_addr,
                "lat": end_coords[0],
                "lng": end_coords[1],
            },
            "sample_points": len(sample_points),
        }
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
