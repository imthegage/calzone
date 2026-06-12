/**
 * Cloudflare Worker: Calendar Proxy with Timezone Fix
 *
 * Proxies ICS calendar requests to outlook.office365.com and converts
 * Windows timezone names to IANA standard names.
 *
 * Timezone mapping source: Unicode CLDR
 * https://github.com/unicode-org/cldr/blob/main/common/supplemental/windowsZones.xml
 */

import { DateTime } from 'luxon';

const WINDOWS_TO_IANA = {
  "Dateline Standard Time": "Etc/GMT+12",
  "UTC-11": "Etc/GMT+11",
  "Aleutian Standard Time": "America/Adak",
  "Hawaiian Standard Time": "Pacific/Honolulu",
  "Marquesas Standard Time": "Pacific/Marquesas",
  "Alaskan Standard Time": "America/Anchorage",
  "UTC-09": "Etc/GMT+9",
  "Pacific Standard Time (Mexico)": "America/Tijuana",
  "UTC-08": "Etc/GMT+8",
  "Pacific Standard Time": "America/Los_Angeles",
  "US Mountain Standard Time": "America/Phoenix",
  "Mountain Standard Time (Mexico)": "America/Mazatlan",
  "Mountain Standard Time": "America/Denver",
  "Yukon Standard Time": "America/Whitehorse",
  "Central America Standard Time": "America/Guatemala",
  "Central Standard Time": "America/Chicago",
  "Easter Island Standard Time": "Pacific/Easter",
  "Central Standard Time (Mexico)": "America/Mexico_City",
  "Canada Central Standard Time": "America/Regina",
  "SA Pacific Standard Time": "America/Bogota",
  "Eastern Standard Time (Mexico)": "America/Cancun",
  "Eastern Standard Time": "America/New_York",
  "Haiti Standard Time": "America/Port-au-Prince",
  "Cuba Standard Time": "America/Havana",
  "US Eastern Standard Time": "America/Indianapolis",
  "Turks And Caicos Standard Time": "America/Grand_Turk",
  "Paraguay Standard Time": "America/Asuncion",
  "Atlantic Standard Time": "America/Halifax",
  "Venezuela Standard Time": "America/Caracas",
  "Central Brazilian Standard Time": "America/Cuiaba",
  "SA Western Standard Time": "America/La_Paz",
  "Pacific SA Standard Time": "America/Santiago",
  "Newfoundland Standard Time": "America/St_Johns",
  "Tocantins Standard Time": "America/Araguaina",
  "E. South America Standard Time": "America/Sao_Paulo",
  "SA Eastern Standard Time": "America/Cayenne",
  "Argentina Standard Time": "America/Buenos_Aires",
  "Greenland Standard Time": "America/Godthab",
  "Montevideo Standard Time": "America/Montevideo",
  "Magallanes Standard Time": "America/Punta_Arenas",
  "Saint Pierre Standard Time": "America/Miquelon",
  "Bahia Standard Time": "America/Bahia",
  "UTC-02": "Etc/GMT+2",
  "Azores Standard Time": "Atlantic/Azores",
  "Cape Verde Standard Time": "Atlantic/Cape_Verde",
  "UTC": "Etc/UTC",
  "GMT Standard Time": "Europe/London",
  "Greenwich Standard Time": "Atlantic/Reykjavik",
  "Sao Tome Standard Time": "Africa/Sao_Tome",
  "Morocco Standard Time": "Africa/Casablanca",
  "W. Europe Standard Time": "Europe/Berlin",
  "Central Europe Standard Time": "Europe/Budapest",
  "Romance Standard Time": "Europe/Paris",
  "Central European Standard Time": "Europe/Warsaw",
  "W. Central Africa Standard Time": "Africa/Lagos",
  "Jordan Standard Time": "Asia/Amman",
  "GTB Standard Time": "Europe/Bucharest",
  "Middle East Standard Time": "Asia/Beirut",
  "Egypt Standard Time": "Africa/Cairo",
  "E. Europe Standard Time": "Europe/Chisinau",
  "Syria Standard Time": "Asia/Damascus",
  "West Bank Standard Time": "Asia/Hebron",
  "South Africa Standard Time": "Africa/Johannesburg",
  "FLE Standard Time": "Europe/Kiev",
  "Israel Standard Time": "Asia/Jerusalem",
  "South Sudan Standard Time": "Africa/Juba",
  "Kaliningrad Standard Time": "Europe/Kaliningrad",
  "Sudan Standard Time": "Africa/Khartoum",
  "Libya Standard Time": "Africa/Tripoli",
  "Namibia Standard Time": "Africa/Windhoek",
  "Arabic Standard Time": "Asia/Baghdad",
  "Turkey Standard Time": "Europe/Istanbul",
  "Arab Standard Time": "Asia/Riyadh",
  "Belarus Standard Time": "Europe/Minsk",
  "Russian Standard Time": "Europe/Moscow",
  "E. Africa Standard Time": "Africa/Nairobi",
  "Iran Standard Time": "Asia/Tehran",
  "Arabian Standard Time": "Asia/Dubai",
  "Astrakhan Standard Time": "Europe/Astrakhan",
  "Azerbaijan Standard Time": "Asia/Baku",
  "Russia Time Zone 3": "Europe/Samara",
  "Mauritius Standard Time": "Indian/Mauritius",
  "Saratov Standard Time": "Europe/Saratov",
  "Georgian Standard Time": "Asia/Tbilisi",
  "Volgograd Standard Time": "Europe/Volgograd",
  "Caucasus Standard Time": "Asia/Yerevan",
  "Afghanistan Standard Time": "Asia/Kabul",
  "West Asia Standard Time": "Asia/Tashkent",
  "Ekaterinburg Standard Time": "Asia/Yekaterinburg",
  "Pakistan Standard Time": "Asia/Karachi",
  "Qyzylorda Standard Time": "Asia/Qyzylorda",
  "India Standard Time": "Asia/Calcutta",
  "Sri Lanka Standard Time": "Asia/Colombo",
  "Nepal Standard Time": "Asia/Katmandu",
  "Central Asia Standard Time": "Asia/Bishkek",
  "Bangladesh Standard Time": "Asia/Dhaka",
  "Omsk Standard Time": "Asia/Omsk",
  "Myanmar Standard Time": "Asia/Rangoon",
  "SE Asia Standard Time": "Asia/Bangkok",
  "Altai Standard Time": "Asia/Barnaul",
  "W. Mongolia Standard Time": "Asia/Hovd",
  "North Asia Standard Time": "Asia/Krasnoyarsk",
  "N. Central Asia Standard Time": "Asia/Novosibirsk",
  "Tomsk Standard Time": "Asia/Tomsk",
  "China Standard Time": "Asia/Shanghai",
  "North Asia East Standard Time": "Asia/Irkutsk",
  "Singapore Standard Time": "Asia/Singapore",
  "W. Australia Standard Time": "Australia/Perth",
  "Taipei Standard Time": "Asia/Taipei",
  "Ulaanbaatar Standard Time": "Asia/Ulaanbaatar",
  "Aus Central W. Standard Time": "Australia/Eucla",
  "Transbaikal Standard Time": "Asia/Chita",
  "Tokyo Standard Time": "Asia/Tokyo",
  "North Korea Standard Time": "Asia/Pyongyang",
  "Korea Standard Time": "Asia/Seoul",
  "Yakutsk Standard Time": "Asia/Yakutsk",
  "Cen. Australia Standard Time": "Australia/Adelaide",
  "AUS Central Standard Time": "Australia/Darwin",
  "E. Australia Standard Time": "Australia/Brisbane",
  "AUS Eastern Standard Time": "Australia/Sydney",
  "West Pacific Standard Time": "Pacific/Port_Moresby",
  "Tasmania Standard Time": "Australia/Hobart",
  "Vladivostok Standard Time": "Asia/Vladivostok",
  "Lord Howe Standard Time": "Australia/Lord_Howe",
  "Bougainville Standard Time": "Pacific/Bougainville",
  "Russia Time Zone 10": "Asia/Srednekolymsk",
  "Magadan Standard Time": "Asia/Magadan",
  "Norfolk Standard Time": "Pacific/Norfolk",
  "Sakhalin Standard Time": "Asia/Sakhalin",
  "Central Pacific Standard Time": "Pacific/Guadalcanal",
  "Russia Time Zone 11": "Asia/Kamchatka",
  "New Zealand Standard Time": "Pacific/Auckland",
  "UTC+12": "Etc/GMT-12",
  "Fiji Standard Time": "Pacific/Fiji",
  "Chatham Islands Standard Time": "Pacific/Chatham",
  "UTC+13": "Etc/GMT-13",
  "Tonga Standard Time": "Pacific/Tongatapu",
  "Samoa Standard Time": "Pacific/Apia",
  "Line Islands Standard Time": "Pacific/Kiritimati",
};

const LA_VTIMEZONE = [
  "BEGIN:VTIMEZONE",
  "TZID:America/Los_Angeles",
  "X-LIC-LOCATION:America/Los_Angeles",
  "BEGIN:DAYLIGHT",
  "TZOFFSETFROM:-0800",
  "TZOFFSETTO:-0700",
  "TZNAME:PDT",
  "DTSTART:19700308T020000",
  "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU",
  "END:DAYLIGHT",
  "BEGIN:STANDARD",
  "TZOFFSETFROM:-0700",
  "TZOFFSETTO:-0800",
  "TZNAME:PST",
  "DTSTART:19701101T020000",
  "RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU",
  "END:STANDARD",
  "END:VTIMEZONE"
].join("\r\n");

function fixTimezones(icsContent) {
  let result = icsContent;

  for (const [windowsTz, ianaTz] of Object.entries(WINDOWS_TO_IANA)) {
    const escaped = windowsTz.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    result = result.replace(
      new RegExp(`(TZID[=:])${escaped}`, "g"),
      `$1${ianaTz}`
    );
  }

  return result;
}

function isValidTimezone(tz) {
  try {
    return !!DateTime.now().setZone(tz).isValid;
  } catch {
    return false;
  }
}

function formatIcsDateTime(dt) {
  return dt.toFormat("yyyyMMdd'T'HHmmss");
}

function convertAllEventTimes(icsContent, targetTz) {
  const eventPattern =
    /^(DTSTART|DTEND|RECURRENCE-ID|EXDATE|RDATE)((?:;[^:]*)?):(\d{8}T\d{6}Z|\d{8}T\d{6})$/gm;

  return icsContent.replace(eventPattern, (match, propName, params = "", rawValue) => {
    let cleanParams = params.replace(/;VALUE=DATE-TIME/gi, "");
    const tzidMatch = cleanParams.match(/;TZID=([^;:]+)/i);
    const sourceTz = tzidMatch ? tzidMatch[1] : null;
    cleanParams = cleanParams.replace(/;TZID=[^;:]*/gi, "");

    let dt;

    if (rawValue.endsWith("Z")) {
      dt = DateTime.fromFormat(rawValue, "yyyyMMdd'T'HHmmss'Z'", { zone: "utc" });
    } else if (sourceTz && isValidTimezone(sourceTz)) {
      dt = DateTime.fromFormat(rawValue, "yyyyMMdd'T'HHmmss", { zone: sourceTz });
    } else {
      dt = DateTime.fromFormat(rawValue, "yyyyMMdd'T'HHmmss", { zone: targetTz });
    }

    if (!dt.isValid) {
      return match;
    }

    const converted = dt.setZone(targetTz);
    return `${propName}${cleanParams};TZID=${targetTz}:${formatIcsDateTime(converted)}`;
  });
}

function ensureVtimezone(icsContent, targetTz) {
  if (targetTz !== "America/Los_Angeles") {
    return icsContent;
  }

  if (/BEGIN:VTIMEZONE[\s\S]*?TZID:America\/Los_Angeles[\s\S]*?END:VTIMEZONE/m.test(icsContent)) {
    return icsContent;
  }

  return icsContent.replace(
    /BEGIN:VCALENDAR\r?\n/,
    `BEGIN:VCALENDAR\r\n${LA_VTIMEZONE}\r\n`
  );
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (!url.pathname.endsWith(".ics")) {
      return env.ASSETS.fetch(request);
    }

    const targetTz = url.searchParams.get("tz") || "America/Los_Angeles";

    if (!isValidTimezone(targetTz)) {
      return new Response(`Invalid timezone: ${targetTz}`, { status: 400 });
    }

    url.searchParams.delete("tz");
    const targetUrl = `https://${env.TARGET_HOST}${url.pathname}${url.search}`;

    try {
      const response = await fetch(targetUrl, {
        headers: { "User-Agent": "Calzone/1.0" },
      });

      if (!response.ok) {
        return new Response(`Upstream error: ${response.status}`, {
          status: response.status,
        });
      }

      const icsContent = await response.text();

      let fixedContent = fixTimezones(icsContent);
      fixedContent = convertAllEventTimes(fixedContent, targetTz);
      fixedContent = ensureVtimezone(fixedContent, targetTz);

      return new Response(fixedContent, {
        headers: {
          "Content-Type": "text/calendar; charset=utf-8",
          "Cache-Control": "public, max-age=3600",
          "Access-Control-Allow-Origin": "*",
        },
      });
    } catch (error) {
      return new Response(`Error: ${error.message}`, { status: 500 });
    }
  },
};
