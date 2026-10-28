/**
 * GeoService.gs
 * Server-Side Geolocation Validation & Haversine Distance Calculation
 * Mother Hub Geofenced Attendance System
 */

const GeoService = {
  EARTH_RADIUS_METERS: 6371000,

  /**
   * Calculates the great-circle distance between two points in decimal degrees
   * using the Haversine formula.
   * 
   * @param {number} lat1 - Latitude of Point 1
   * @param {number} lon1 - Longitude of Point 1
   * @param {number} lat2 - Latitude of Point 2
   * @param {number} lon2 - Longitude of Point 2
   * @return {number} Distance in meters rounded to 2 decimal places
   */
  haversineDistance: function(lat1, lon1, lat2, lon2) {
    if (lat1 === null || lon1 === null || lat2 === null || lon2 === null ||
        isNaN(lat1) || isNaN(lon1) || isNaN(lat2) || isNaN(lon2)) {
      throw new Error("Invalid coordinate parameters provided for distance calculation.");
    }

    const toRad = function(degree) {
      return (degree * Math.PI) / 180;
    };

    const dLat = toRad(lat2 - lat1);
    const dLon = toRad(lon2 - lon1);

    const phi1 = toRad(lat1);
    const phi2 = toRad(lat2);

    const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
              Math.cos(phi1) * Math.cos(phi2) *
              Math.sin(dLon / 2) * Math.sin(dLon / 2);

    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));

    const distanceMeters = this.EARTH_RADIUS_METERS * c;
    return Math.round(distanceMeters * 100) / 100;
  },

  /**
   * Verifies if user coordinates fall strictly within the Mother Hub's designated radius
   * plus an optional server-side tolerance buffer.
   * 
   * @param {number} userLat
   * @param {number} userLng
   * @param {number} hubLat
   * @param {number} hubLng
   * @param {number} radiusMeters
   * @param {number} toleranceBufferMeters
   * @return {{isInside: boolean, distanceMeters: number, maxAllowedRadius: number}}
   */
  verifyGeofence: function(userLat, userLng, hubLat, hubLng, radiusMeters, toleranceBufferMeters) {
    const radius = Number(radiusMeters) || 100;
    const buffer = Number(toleranceBufferMeters) || 0;
    const maxAllowed = radius + buffer;
    
    const distance = this.haversineDistance(userLat, userLng, hubLat, hubLng);
    const isInside = distance <= maxAllowed;

    return {
      isInside: isInside,
      distanceMeters: distance,
      maxAllowedRadius: maxAllowed
    };
  }
};
