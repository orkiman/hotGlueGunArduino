#include <Arduino.h>
#include <ArduinoJson.h>
#include <EEPROM.h>

// Set to 1 to enable Serial debug logs
#define DEBUG 0

#if DEBUG
  #define DBG_PRINTLN(x) Serial.println(x)
  #define DBG_PRINT(x) Serial.print(x)
#else
  #define DBG_PRINTLN(x) do {} while (0)
  #define DBG_PRINT(x) do {} while (0)
#endif

/*
Example NDJSON commands (one JSON object per line):
{"cmd":"set_active","active":true}
{"cmd":"set_config","pulses_per_mm":12.34,"max_ms_per_mm":80,"photocell_offset_mm":250.0,"debounce_ms":20}
{"cmd":"set_pattern","gun":1,"lines":[{"start":10,"end":40},{"start":60,"end":90}]}
{"cmd":"set_pattern","gun":2,"lines":[{"start":15,"end":35}]}
{"cmd":"calib_arm","paper_length_mm":297.0}
{"cmd":"test_open","gun":1,"timeout_ms":2000}
{"cmd":"test_close","gun":1}
*/

// Pins
static constexpr uint8_t PIN_ENCODER_A = 2;
static constexpr uint8_t PIN_PHOTOCELL = 3;
static constexpr uint8_t PIN_GUN1 = 8;
static constexpr uint8_t PIN_GUN2 = 9;

struct Line {
  float start_mm;
  float end_mm;
};

static constexpr uint8_t MAX_LINES_PER_GUN = 32;

struct Pattern {
  Line lines[MAX_LINES_PER_GUN];
  uint8_t count;
};

struct Config {
  float pulses_per_mm;
  uint16_t max_ms_per_mm;
  float photocell_offset_mm;
  uint16_t debounce_ms;
};

struct Sheet {
  bool active;
  float mm;
  uint32_t started_ms;

  int32_t last_mm_int;
  uint32_t last_mm_change_ms;

  bool slow_block[2];
};

static constexpr uint8_t MAX_SHEETS = 10;
static constexpr uint32_t SHEET_TIMEOUT_MS = 30000;
static constexpr float SHEET_END_MARGIN_MM = 30.0f;

static constexpr uint32_t PERSIST_MAGIC = 0x48474743UL; // 'HGGC'
static constexpr uint16_t PERSIST_VERSION = 1;

struct PersistData {
  uint32_t magic;
  uint16_t version;
  Config cfg;
  Pattern pat1;
  Pattern pat2;
  uint16_t crc;
};

static Config g_cfg;
static Pattern g_pat1;
static Pattern g_pat2;
static Sheet g_sheets[MAX_SHEETS];

static bool g_active = false;

static volatile uint32_t g_encoder_pulses_isr = 0;
static uint32_t g_encoder_pulses_total = 0;

static float g_current_mm = 0.0f;

// Photocell debounce
static bool g_photo_raw_last = true;
static bool g_photo_stable = true;
static uint32_t g_photo_last_change_ms = 0;

// Test override
static bool g_test_override_on[2] = {false, false};
static uint32_t g_test_override_until_ms[2] = {0, 0};

// Calibration
static bool g_calib_armed = false;
static bool g_calib_waiting_for_second = false;
static float g_calib_paper_length_mm = 0.0f;
static uint32_t g_calib_pulse_start = 0;

static uint16_t crc16_update(uint16_t crc, uint8_t data) {
  crc ^= data;
  for (uint8_t i = 0; i < 8; i++) {
    if (crc & 1) {
      crc = (crc >> 1) ^ 0xA001;
    } else {
      crc >>= 1;
    }
  }
  return crc;
}

static uint16_t crc16(const uint8_t *data, size_t len) {
  uint16_t crc = 0xFFFF;
  for (size_t i = 0; i < len; i++) {
    crc = crc16_update(crc, data[i]);
  }
  return crc;
}

static uint16_t computePersistCrc(const PersistData &pd) {
  PersistData tmp = pd;
  tmp.crc = 0;
  return crc16(reinterpret_cast<const uint8_t *>(&tmp), sizeof(PersistData));
}

static void defaultConfig() {
  g_cfg.pulses_per_mm = 10.0f;
  g_cfg.max_ms_per_mm = 100;
  g_cfg.photocell_offset_mm = 0.0f;
  g_cfg.debounce_ms = 20;
}

static void clearPatterns() {
  g_pat1.count = 0;
  g_pat2.count = 0;
}

static void clearSheets() {
  for (uint8_t i = 0; i < MAX_SHEETS; i++) {
    g_sheets[i].active = false;
  }
}

static float patternEndMm() {
  float end_mm = 0.0f;
  for (uint8_t i = 0; i < g_pat1.count; i++) {
    end_mm = max(end_mm, max(g_pat1.lines[i].start_mm, g_pat1.lines[i].end_mm));
  }
  for (uint8_t i = 0; i < g_pat2.count; i++) {
    end_mm = max(end_mm, max(g_pat2.lines[i].start_mm, g_pat2.lines[i].end_mm));
  }
  return end_mm;
}

static void loadConfig() {
  defaultConfig();
  clearPatterns();

  if (EEPROM.length() < (int)sizeof(PersistData)) {
    DBG_PRINTLN("EEPROM too small for PersistData; using defaults");
    return;
  }

  PersistData pd;
  EEPROM.get(0, pd);
  if (pd.magic != PERSIST_MAGIC || pd.version != PERSIST_VERSION) {
    DBG_PRINTLN("No valid persisted data; using defaults");
    return;
  }

  uint16_t expected = computePersistCrc(pd);
  if (expected != pd.crc) {
    DBG_PRINTLN("Persist CRC mismatch; using defaults");
    return;
  }

  g_cfg = pd.cfg;
  g_pat1 = pd.pat1;
  g_pat2 = pd.pat2;
}

static void saveConfig() {
  if (EEPROM.length() < (int)sizeof(PersistData)) {
    return;
  }

  PersistData pd;
  pd.magic = PERSIST_MAGIC;
  pd.version = PERSIST_VERSION;
  pd.cfg = g_cfg;
  pd.pat1 = g_pat1;
  pd.pat2 = g_pat2;
  pd.crc = computePersistCrc(pd);

  EEPROM.put(0, pd);

  #if defined(ESP8266) || defined(ESP32)
    EEPROM.commit();
  #endif
}

static void encoderISR() {
  g_encoder_pulses_isr++;
}

static void setGunOutputs(bool gun1_on, bool gun2_on) {
  digitalWrite(PIN_GUN1, gun1_on ? HIGH : LOW);
  digitalWrite(PIN_GUN2, gun2_on ? HIGH : LOW);
}

static bool photocellFallingEdgeDebounced() {
  bool raw = (digitalRead(PIN_PHOTOCELL) == HIGH);
  uint32_t now = millis();

  if (raw != g_photo_raw_last) {
    g_photo_raw_last = raw;
    g_photo_last_change_ms = now;
  }

  if ((uint32_t)(now - g_photo_last_change_ms) >= g_cfg.debounce_ms) {
    if (g_photo_stable != raw) {
      bool prev = g_photo_stable;
      g_photo_stable = raw;
      if (prev == true && g_photo_stable == false) {
        return true;
      }
    }
  }

  return false;
}

static int8_t parseGunSelector(JsonVariant v) {
  if (v.is<const char *>()) {
    const char *s = v.as<const char *>();
    if (!s) return -1;
    if (strcmp(s, "both") == 0) return 3;
  }
  if (v.is<int>()) {
    int g = v.as<int>();
    if (g == 1) return 1;
    if (g == 2) return 2;
  }
  return -1;
}

static void handleTestOpen(int8_t which, uint32_t timeout_ms) {
  uint32_t now = millis();
  uint32_t until = now + timeout_ms;

  if (which == 1 || which == 3) {
    g_test_override_on[0] = true;
    g_test_override_until_ms[0] = until;
  }
  if (which == 2 || which == 3) {
    g_test_override_on[1] = true;
    g_test_override_until_ms[1] = until;
  }
}

static void handleTestClose(int8_t which) {
  if (which == 1 || which == 3) {
    g_test_override_on[0] = false;
  }
  if (which == 2 || which == 3) {
    g_test_override_on[1] = false;
  }
}

static void serviceTestOverrides() {
  uint32_t now = millis();
  for (uint8_t i = 0; i < 2; i++) {
    if (g_test_override_on[i]) {
      if ((int32_t)(now - g_test_override_until_ms[i]) >= 0) {
        g_test_override_on[i] = false;
      }
    }
  }
}

static void addSheet() {
  uint32_t now = millis();

  int best = -1;
  for (uint8_t i = 0; i < MAX_SHEETS; i++) {
    if (!g_sheets[i].active) {
      best = i;
      break;
    }
  }

  if (best < 0) {
    uint32_t oldest_ms = 0xFFFFFFFFUL;
    for (uint8_t i = 0; i < MAX_SHEETS; i++) {
      if (g_sheets[i].active && g_sheets[i].started_ms < oldest_ms) {
        oldest_ms = g_sheets[i].started_ms;
        best = i;
      }
    }
  }

  if (best < 0) return;

  Sheet &s = g_sheets[best];
  s.active = true;
  s.mm = -g_cfg.photocell_offset_mm;
  s.started_ms = now;

  s.last_mm_int = (int32_t)floor(s.mm);
  s.last_mm_change_ms = now;
  s.slow_block[0] = false;
  s.slow_block[1] = false;
}

static bool mmInAnyInterval(const Pattern &pat, float mm) {
  for (uint8_t i = 0; i < pat.count; i++) {
    float a = pat.lines[i].start_mm;
    float b = pat.lines[i].end_mm;
    float lo = min(a, b);
    float hi = max(a, b);
    if (mm >= lo && mm <= hi) {
      return true;
    }
  }
  return false;
}

static void updateSheetsByDelta(float delta_mm) {
  uint32_t now = millis();
  float end_mm = patternEndMm() + SHEET_END_MARGIN_MM;

  for (uint8_t i = 0; i < MAX_SHEETS; i++) {
    Sheet &s = g_sheets[i];
    if (!s.active) continue;

    s.mm += delta_mm;

    int32_t mm_int = (int32_t)floor(s.mm);
    if (mm_int != s.last_mm_int) {
      uint32_t step_ms = (uint32_t)(now - s.last_mm_change_ms);
      s.last_mm_int = mm_int;
      s.last_mm_change_ms = now;

      for (uint8_t g = 0; g < 2; g++) {
        if (step_ms > g_cfg.max_ms_per_mm) {
          s.slow_block[g] = true;
        } else {
          s.slow_block[g] = false;
        }
      }
    } else {
      for (uint8_t g = 0; g < 2; g++) {
        if ((uint32_t)(now - s.last_mm_change_ms) > g_cfg.max_ms_per_mm) {
          s.slow_block[g] = true;
        }
      }
    }

    if (s.mm > end_mm) {
      s.active = false;
      continue;
    }

    if ((uint32_t)(now - s.started_ms) > SHEET_TIMEOUT_MS) {
      s.active = false;
      continue;
    }
  }
}

static void evaluateOutputs(bool &gun1_on, bool &gun2_on) {
  gun1_on = false;
  gun2_on = false;

  for (uint8_t i = 0; i < MAX_SHEETS; i++) {
    const Sheet &s = g_sheets[i];
    if (!s.active) continue;

    bool req1 = mmInAnyInterval(g_pat1, s.mm);
    bool req2 = mmInAnyInterval(g_pat2, s.mm);

    if (req1 && !s.slow_block[0]) gun1_on = true;
    if (req2 && !s.slow_block[1]) gun2_on = true;

    if (gun1_on && gun2_on) {
      return;
    }
  }
}

static void handlePhotocellEvent() {
  if (g_calib_armed) {
    if (!g_calib_waiting_for_second) {
      g_calib_pulse_start = g_encoder_pulses_total;
      g_calib_waiting_for_second = true;
      return;
    } else {
      uint32_t measured = g_encoder_pulses_total - g_calib_pulse_start;
      if (g_calib_paper_length_mm > 0.001f) {
        g_cfg.pulses_per_mm = (float)measured / g_calib_paper_length_mm;
        saveConfig();

        StaticJsonDocument<128> out;
        out["event"] = "calib_result";
        out["pulses_per_mm"] = g_cfg.pulses_per_mm;
        serializeJson(out, Serial);
        Serial.print('\n');
      }

      g_calib_armed = false;
      g_calib_waiting_for_second = false;
      return;
    }
  }

  if (!g_active) {
    return;
  }

  addSheet();
  g_current_mm = -g_cfg.photocell_offset_mm;
}

static void applySetConfig(JsonDocument &doc) {
  bool changed = false;

  if (doc.containsKey("pulses_per_mm")) {
    float v = doc["pulses_per_mm"].as<float>();
    if (v > 0.0001f && v != g_cfg.pulses_per_mm) {
      g_cfg.pulses_per_mm = v;
      changed = true;
    }
  }
  if (doc.containsKey("max_ms_per_mm")) {
    int v = doc["max_ms_per_mm"].as<int>();
    if (v >= 1 && v <= 60000 && (uint16_t)v != g_cfg.max_ms_per_mm) {
      g_cfg.max_ms_per_mm = (uint16_t)v;
      changed = true;
    }
  }
  if (doc.containsKey("photocell_offset_mm")) {
    float v = doc["photocell_offset_mm"].as<float>();
    if (v != g_cfg.photocell_offset_mm) {
      g_cfg.photocell_offset_mm = v;
      changed = true;
    }
  }
  if (doc.containsKey("debounce_ms")) {
    int v = doc["debounce_ms"].as<int>();
    if (v >= 0 && v <= 1000 && (uint16_t)v != g_cfg.debounce_ms) {
      g_cfg.debounce_ms = (uint16_t)v;
      changed = true;
    }
  }

  if (changed) {
    saveConfig();
  }
}

static void applySetPattern(JsonDocument &doc) {
  int gun = doc["gun"].as<int>();
  JsonArray lines = doc["lines"].as<JsonArray>();
  if (gun != 1 && gun != 2) return;
  if (lines.isNull()) return;

  Pattern &pat = (gun == 1) ? g_pat1 : g_pat2;
  pat.count = 0;

  for (JsonObject lineObj : lines) {
    if (pat.count >= MAX_LINES_PER_GUN) break;
    if (!lineObj.containsKey("start") || !lineObj.containsKey("end")) continue;

    pat.lines[pat.count].start_mm = lineObj["start"].as<float>();
    pat.lines[pat.count].end_mm = lineObj["end"].as<float>();
    pat.count++;
  }

  saveConfig();
}

static void handleJsonLine(const char *line) {
  if (!line || !line[0]) return;

  StaticJsonDocument<2048> doc;
  DeserializationError err = deserializeJson(doc, line);
  if (err) {
    return;
  }

  const char *cmd = doc["cmd"].as<const char *>();
  if (!cmd) return;

  if (strcmp(cmd, "set_active") == 0) {
    bool a = doc["active"].as<bool>();
    g_active = a;
    if (!g_active) {
      clearSheets();
      g_test_override_on[0] = false;
      g_test_override_on[1] = false;
      setGunOutputs(false, false);
    }
    return;
  }

  if (strcmp(cmd, "set_config") == 0) {
    applySetConfig(doc);
    return;
  }

  if (strcmp(cmd, "set_pattern") == 0) {
    applySetPattern(doc);
    return;
  }

  if (strcmp(cmd, "calib_arm") == 0) {
    float len = doc["paper_length_mm"].as<float>();
    if (len > 0.001f) {
      g_calib_paper_length_mm = len;
      g_calib_armed = true;
      g_calib_waiting_for_second = false;
    }
    return;
  }

  if (strcmp(cmd, "test_open") == 0) {
    int8_t which = parseGunSelector(doc["gun"]);
    uint32_t timeout_ms = doc.containsKey("timeout_ms") ? doc["timeout_ms"].as<uint32_t>() : 2000;
    if (timeout_ms < 1) timeout_ms = 1;
    if (timeout_ms > 600000UL) timeout_ms = 600000UL;
    if (which > 0) {
      handleTestOpen(which, timeout_ms);
    }
    return;
  }

  if (strcmp(cmd, "test_close") == 0) {
    int8_t which = parseGunSelector(doc["gun"]);
    if (which > 0) {
      handleTestClose(which);
    }
    return;
  }
}

static void serviceSerialJson() {
  static char buf[512];
  static size_t idx = 0;

  while (Serial.available() > 0) {
    int c = Serial.read();
    if (c < 0) break;

    if (c == '\r') {
      continue;
    }

    if (c == '\n') {
      buf[idx] = '\0';
      if (idx > 0) {
        handleJsonLine(buf);
      }
      idx = 0;
      continue;
    }

    if (idx < sizeof(buf) - 1) {
      buf[idx++] = (char)c;
    } else {
      idx = 0;
    }
  }
}

void setup() {
  pinMode(PIN_ENCODER_A, INPUT_PULLUP);
  pinMode(PIN_PHOTOCELL, INPUT_PULLUP);

  pinMode(PIN_GUN1, OUTPUT);
  pinMode(PIN_GUN2, OUTPUT);
  setGunOutputs(false, false);

  Serial.begin(115200);

  loadConfig();
  clearSheets();

  g_photo_raw_last = (digitalRead(PIN_PHOTOCELL) == HIGH);
  g_photo_stable = g_photo_raw_last;
  g_photo_last_change_ms = millis();

  attachInterrupt(digitalPinToInterrupt(PIN_ENCODER_A), encoderISR, RISING);
}

void loop() {
  serviceSerialJson();
  serviceTestOverrides();

  uint32_t pulses;
  noInterrupts();
  pulses = g_encoder_pulses_isr;
  g_encoder_pulses_isr = 0;
  interrupts();

  if (pulses > 0) {
    g_encoder_pulses_total += pulses;

    float ppm = g_cfg.pulses_per_mm;
    if (ppm < 0.0001f) ppm = 0.0001f;
    float delta_mm = (float)pulses / ppm;

    g_current_mm += delta_mm;
    updateSheetsByDelta(delta_mm);
  } else {
    updateSheetsByDelta(0.0f);
  }

  if (photocellFallingEdgeDebounced()) {
    handlePhotocellEvent();
  }

  bool gun1_req = false;
  bool gun2_req = false;
  if (g_active) {
    evaluateOutputs(gun1_req, gun2_req);
  }

  bool gun1_on = gun1_req;
  bool gun2_on = gun2_req;

  if (g_test_override_on[0]) gun1_on = true;
  if (g_test_override_on[1]) gun2_on = true;

  if (!g_active && !g_test_override_on[0] && !g_test_override_on[1]) {
    gun1_on = false;
    gun2_on = false;
  }

  setGunOutputs(gun1_on, gun2_on);
}
