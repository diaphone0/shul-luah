"""
NOAA solar position calculator.

Ported from mod_NOAACalculator.bas in https://github.com/diaphone1/vbzmanim
(itself ported from the NOAA Solar Calculator algorithm, via KosherJava/Zmanim).

Computes UTC sunrise/sunset times given a Julian day, geographic location, and
solar zenith angle. This is the core astronomical engine underlying all zmanim
(halachic times) calculations in zmanim.py.
"""
from __future__ import annotations

import math
from dataclasses import dataclass

REFRACTION = 34.0 / 60.0
SOLAR_RADIUS = 16.0 / 60.0
EARTH_RADIUS = 6356.9  # km


@dataclass
class Location:
    """Geographic location: latitude/longitude in degrees, elevation in meters."""
    latitude: float
    longitude: float
    elevation: float = 0.0


def deg_to_rad(angle_deg: float) -> float:
    return math.radians(angle_deg)


def rad_to_deg(angle_rad: float) -> float:
    return math.degrees(angle_rad)


def calc_time_julian_cent(jd: float) -> float:
    return (jd - 2451545.0) / 36525.0


def calc_jd_from_julian_cent(jcent: float) -> float:
    return jcent * 36525.0 + 2451545.0


def calc_geom_mean_long_sun(jcent: float) -> float:
    gmls = 280.46646 + jcent * (36000.76983 + 0.0003032 * jcent)
    return gmls % 360.0


def calc_geom_mean_anomaly_sun(jcent: float) -> float:
    return 357.52911 + jcent * (35999.05029 - 0.0001537 * jcent)


def calc_eccentricity_earth_orbit(jcent: float) -> float:
    return 0.016708634 - jcent * (0.000042037 + 0.0000001267 * jcent)


def calc_sun_eq_of_center(jcent: float) -> float:
    m = calc_geom_mean_anomaly_sun(jcent)
    mrad = deg_to_rad(m)
    sinm = math.sin(mrad)
    sin2m = math.sin(mrad + mrad)
    sin3m = math.sin(mrad + mrad + mrad)
    return (
        sinm * (1.914602 - jcent * (0.004817 + 0.000014 * jcent))
        + sin2m * (0.019993 - 0.000101 * jcent)
        + sin3m * 0.000289
    )


def calc_sun_true_long(jcent: float) -> float:
    return calc_geom_mean_long_sun(jcent) + calc_sun_eq_of_center(jcent)


def calc_sun_apparent_long(jcent: float) -> float:
    stl = calc_sun_true_long(jcent)
    omega = 125.04 - 1934.136 * jcent
    return stl - 0.00569 - 0.00478 * math.sin(deg_to_rad(omega))


def calc_mean_obliquity_of_ecliptic(jcent: float) -> float:
    seconds = 21.448 - jcent * (46.815 + jcent * (0.00059 - jcent * 0.001813))
    return 23.0 + (26.0 + (seconds / 60.0)) / 60.0


def calc_obliquity_correction(jcent: float) -> float:
    mooe = calc_mean_obliquity_of_ecliptic(jcent)
    omega = 125.04 - 1934.136 * jcent
    return mooe + 0.00256 * math.cos(deg_to_rad(omega))


def calc_sun_declination(jcent: float) -> float:
    oc = calc_obliquity_correction(jcent)
    sal = calc_sun_apparent_long(jcent)
    sint = math.sin(deg_to_rad(oc)) * math.sin(deg_to_rad(sal))
    return rad_to_deg(math.asin(sint))


def calc_equation_of_time(jcent: float) -> float:
    oc = calc_obliquity_correction(jcent)
    gmls = calc_geom_mean_long_sun(jcent)
    eeo = calc_eccentricity_earth_orbit(jcent)
    gmas = calc_geom_mean_anomaly_sun(jcent)

    y = math.tan(deg_to_rad(oc) / 2.0)
    y = y * y

    sin2gmls = math.sin(2.0 * deg_to_rad(gmls))
    singmas = math.sin(deg_to_rad(gmas))
    cos2gmls = math.cos(2.0 * deg_to_rad(gmls))
    sin4gmls = math.sin(4.0 * deg_to_rad(gmls))
    sin2gmas = math.sin(2.0 * deg_to_rad(gmas))

    etime = (
        y * sin2gmls
        - 2.0 * eeo * singmas
        + 4.0 * eeo * y * singmas * cos2gmls
        - 0.5 * y * y * sin4gmls
        - 1.25 * eeo * eeo * sin2gmas
    )
    return rad_to_deg(etime) * 4.0


def calc_hour_angle_sunrise(lat: float, solar_dec: float, zenith: float) -> float:
    lat_rad = deg_to_rad(lat)
    sd_rad = deg_to_rad(solar_dec)
    return math.acos(
        math.cos(deg_to_rad(zenith)) / (math.cos(lat_rad) * math.cos(sd_rad))
        - math.tan(lat_rad) * math.tan(sd_rad)
    )


def calc_hour_angle_sunset(lat: float, solar_dec: float, zenith: float) -> float:
    return -calc_hour_angle_sunrise(lat, solar_dec, zenith)


def calc_sol_noon_utc(jd: float, longitude: float) -> float:
    jcent = calc_time_julian_cent(jd)

    tnoon = calc_time_julian_cent(calc_jd_from_julian_cent(jcent) + longitude / 360.0)
    eq_time = calc_equation_of_time(tnoon)
    sol_noon_utc = 720 + (longitude * 4) - eq_time

    newt = calc_time_julian_cent(
        calc_jd_from_julian_cent(jcent) - 0.5 + sol_noon_utc / 1440.0
    )
    eq_time = calc_equation_of_time(newt)
    sol_noon_utc = 720 + (longitude * 4) - eq_time

    return sol_noon_utc


def calc_sunrise_utc(jd: float, latitude: float, longitude: float, zenith: float) -> float:
    jcent = calc_time_julian_cent(jd)

    noonmin = calc_sol_noon_utc(jd, longitude)
    tnoon = calc_time_julian_cent(jd + noonmin / 1440.0)

    eq_time = calc_equation_of_time(tnoon)
    solar_dec = calc_sun_declination(tnoon)
    hour_angle = calc_hour_angle_sunrise(latitude, solar_dec, zenith)

    delta = longitude - rad_to_deg(hour_angle)
    time_diff = 4 * delta
    time_utc = 720 + time_diff - eq_time

    newt = calc_time_julian_cent(calc_jd_from_julian_cent(jcent) + time_utc / 1440.0)
    eq_time = calc_equation_of_time(newt)
    solar_dec = calc_sun_declination(newt)
    hour_angle = calc_hour_angle_sunrise(latitude, solar_dec, zenith)
    delta = longitude - rad_to_deg(hour_angle)
    time_diff = 4 * delta
    time_utc = 720 + time_diff - eq_time

    return time_utc


def calc_sunset_utc(jd: float, latitude: float, longitude: float, zenith: float) -> float:
    jcent = calc_time_julian_cent(jd)

    noonmin = calc_sol_noon_utc(jd, longitude)
    tnoon = calc_time_julian_cent(jd + noonmin / 1440.0)

    eq_time = calc_equation_of_time(tnoon)
    solar_dec = calc_sun_declination(tnoon)
    hour_angle = calc_hour_angle_sunset(latitude, solar_dec, zenith)

    delta = longitude - rad_to_deg(hour_angle)
    time_diff = 4 * delta
    time_utc = 720 + time_diff - eq_time

    newt = calc_time_julian_cent(calc_jd_from_julian_cent(jcent) + time_utc / 1440.0)
    eq_time = calc_equation_of_time(newt)
    solar_dec = calc_sun_declination(newt)
    hour_angle = calc_hour_angle_sunset(latitude, solar_dec, zenith)

    delta = longitude - rad_to_deg(hour_angle)
    time_diff = 4 * delta
    time_utc = 720 + time_diff - eq_time

    return time_utc


def get_elevation_adjustment(elevation: float) -> float:
    return rad_to_deg(math.acos(EARTH_RADIUS / (EARTH_RADIUS + (elevation / 1000))))


def adjust_zenith(zenith: float, elevation: float) -> float:
    if zenith == 90.0:
        return zenith + (SOLAR_RADIUS + REFRACTION + get_elevation_adjustment(elevation))
    return zenith


def get_utc_sunrise(
    jd: float, here: Location, zenith: float, adjust_for_elevation: bool
) -> float:
    """Returns fractional UTC hour (0-24) of sunrise for the given Julian day."""
    elevation = here.elevation if adjust_for_elevation else 0.0
    adjusted_zenith = adjust_zenith(zenith, elevation)

    sunrise = calc_sunrise_utc(jd, here.latitude, -here.longitude, adjusted_zenith)
    sunrise = sunrise / 60
    sunrise %= 24.0
    return sunrise


def get_utc_sunset(
    jd: float, here: Location, zenith: float, adjust_for_elevation: bool
) -> float:
    """Returns fractional UTC hour (0-24) of sunset for the given Julian day."""
    elevation = here.elevation if adjust_for_elevation else 0.0
    adjusted_zenith = adjust_zenith(zenith, elevation)

    sunset = calc_sunset_utc(jd, here.latitude, -here.longitude, adjusted_zenith)
    sunset = sunset / 60
    sunset %= 24.0
    return sunset
