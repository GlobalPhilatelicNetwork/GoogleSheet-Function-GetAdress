/**
 * Ruft eine Client-Adresse von der Heinrich Köhler API ab.
 *
 * @param {string} user - API-Benutzername für Basic Auth
 * @param {string} pw - API-Passwort für Basic Auth
 * @param {number} nid - Client Node-ID
 * @param {string} [type] - Adresstyp (z.B. "shipping_for_lots", "website"). Leer = Defaultadresse.
 * @param {string} [feld] - Einzelnes Feld (z.B. "locality", "postal_code"). Leer = gerenderte Adresse.
 * @return {string} Die aufbereitete Adresse oder das angeforderte Feld
 * @customfunction
 */
function GetBOAdress(user, pw, nid, type, feld) {
  if (!user || !pw || !nid) {
    throw new Error("User, PW und NID sind Pflichtparameter.");
  }

  var url = "https://heinrich-koehler.de/en/api/client-address?filter[client]=" + encodeURIComponent(nid);
  var headers = {
    "Authorization": "Basic " + Utilities.base64Encode(user + ":" + pw)
  };

  var response = UrlFetchApp.fetch(url, {
    headers: headers,
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  if (code !== 200) {
    throw new Error("API-Fehler: HTTP " + code);
  }

  var json = JSON.parse(response.getContentText());
  var addresses = json.data;

  if (!addresses || addresses.length === 0) {
    throw new Error("Keine Adressen für NID " + nid + " gefunden.");
  }

  // Defaultadresse finden
  var defaultAddress = null;
  for (var i = 0; i < addresses.length; i++) {
    if (addresses[i].default_address === true) {
      defaultAddress = addresses[i];
      break;
    }
  }
  // Fallback: erste Adresse, wenn keine als Default markiert
  if (!defaultAddress) {
    defaultAddress = addresses[0];
  }

  // Zieladresse bestimmen
  var targetAddress = defaultAddress;

  if (type && type !== "") {
    var typeLower = type.toLowerCase().trim();
    for (var j = 0; j < addresses.length; j++) {
      var types = addresses[j].address_types;
      if (types) {
        for (var k = 0; k < types.length; k++) {
          if (types[k].toLowerCase() === typeLower) {
            targetAddress = addresses[j];
            break;
          }
        }
      }
      if (targetAddress !== defaultAddress) break;
    }
    // Wenn Typ nicht gefunden: Fallback auf Default (targetAddress bleibt defaultAddress)
  }

  // Einzelnes Feld zurückgeben
  if (feld && feld !== "") {
    var feldLower = feld.toLowerCase().trim();
    var addressData = targetAddress.address;
    if (addressData && addressData.hasOwnProperty(feldLower)) {
      var value = addressData[feldLower];
      return value !== null && value !== "" ? String(value) : "";
    }
    // Fallback auf Defaultadresse, falls Feld in Zieladresse nicht vorhanden
    if (targetAddress !== defaultAddress) {
      addressData = defaultAddress.address;
      if (addressData && addressData.hasOwnProperty(feldLower)) {
        var fallbackValue = addressData[feldLower];
        return fallbackValue !== null && fallbackValue !== "" ? String(fallbackValue) : "";
      }
    }
    throw new Error("Unbekanntes Feld: " + feld);
  }

  // Gerenderte Adresse als reinen Text zurückgeben
  var rendered = targetAddress.rendered || "";
  return htmlToPlainText_(rendered);
}

/**
 * Wandelt HTML in reinen Text um.
 * Entfernt alle Tags, dekodiert HTML-Entities und bereinigt Whitespace.
 *
 * @param {string} html - HTML-String
 * @return {string} Reiner Text
 * @private
 */
function htmlToPlainText_(html) {
  if (!html) return "";

  var text = html;

  // <br>, <br/>, </div>, </p> durch Zeilenumbrüche ersetzen
  text = text.replace(/<br\s*\/?>/gi, "\n");
  text = text.replace(/<\/div>/gi, "\n");
  text = text.replace(/<\/p>/gi, "\n");

  // Alle verbleibenden HTML-Tags entfernen
  text = text.replace(/<[^>]+>/g, "");

  // HTML-Entities dekodieren
  text = decodeHtmlEntities_(text);

  // Mehrfache Leerzeichen zu einem reduzieren
  text = text.replace(/[ \t]+/g, " ");

  // Mehrfache Zeilenumbrüche reduzieren und trimmen
  text = text.replace(/\n\s*\n/g, "\n");
  text = text.trim();

  return text;
}

/**
 * Dekodiert gängige HTML-Entities.
 *
 * @param {string} text - Text mit HTML-Entities
 * @return {string} Dekodierter Text
 * @private
 */
function decodeHtmlEntities_(text) {
  var entities = {
    "&amp;": "&",
    "&lt;": "<",
    "&gt;": ">",
    "&quot;": "\"",
    "&#039;": "'",
    "&apos;": "'",
    "&nbsp;": " ",
    "&ndash;": "\u2013",
    "&mdash;": "\u2014",
    "&ouml;": "\u00f6",
    "&uuml;": "\u00fc",
    "&auml;": "\u00e4",
    "&Ouml;": "\u00d6",
    "&Uuml;": "\u00dc",
    "&Auml;": "\u00c4",
    "&szlig;": "\u00df"
  };

  for (var entity in entities) {
    text = text.split(entity).join(entities[entity]);
  }

  // Numerische Entities (&#123; und &#x1A;)
  text = text.replace(/&#(\d+);/g, function(match, dec) {
    return String.fromCharCode(parseInt(dec, 10));
  });
  text = text.replace(/&#x([0-9a-fA-F]+);/g, function(match, hex) {
    return String.fromCharCode(parseInt(hex, 16));
  });

  return text;
}
