#include <WiFi.h>
#include <HTTPClient.h>
#include <Adafruit_Fingerprint.h>

// Your Wi-Fi network credentials
const char* ssid = "TP-Link_1A24_5G";
const char* password = "58367816";

// Replace with your computer's local IP address running Flask (e.g., 192.168.0.X)
const char* serverUrl = "http://192.168.0.111:5000/api/fingerprint-scan";

// Use Hardware Serial 2 on ESP32 (RX=16, TX=17)
HardwareSerial mySerial(2);
Adafruit_Fingerprint finger = Adafruit_Fingerprint(&mySerial);

void setup() {
  Serial.begin(115200); // USB Serial for debugging
  mySerial.begin(57600, SERIAL_8N1, 16, 17); // Sensor serial
  
  // Connect to Wi-Fi
  WiFi.begin(ssid, password);
  Serial.print("Connecting to Wi-Fi");
  while (WiFi.status() != WL_CONNECTED) {
    delay(500);
    Serial.print(".");
  }
  Serial.println("\n--- Wi-Fi Connected ---");
  Serial.print("ESP32 IP Address: "); 
  Serial.println(WiFi.localIP());

  // Initialize Fingerprint Sensor
  finger.begin(57600);
  if (finger.verifyPassword()) {
    Serial.println("Fingerprint sensor found and ready!");
  } else {
    Serial.println("Sensor not found! Check wiring.");
    while (1) { delay(1); }
  }
}

void loop() {
  int fingerID = getFingerprintID();
  if (fingerID >= 0) {
    // 1. Send HTTP POST to Flask backend when matched
    if (WiFi.status() == WL_CONNECTED) {
      HTTPClient http;
      http.begin(serverUrl);
      http.addHeader("Content-Type", "application/json");
      
      String payload = "{\"finger_id\":" + String(fingerID) + "}";
      int httpResponseCode = http.POST(payload);
      
      if (httpResponseCode > 0) {
        Serial.println("Access signal successfully sent to Flask server!");
      } else {
        Serial.print("Error sending POST: "); 
        Serial.println(httpResponseCode);
      }
      http.end();
    }
    
    // 2. Output over USB Serial for wired debugging
    Serial.print("MATCH_FOUND:");
    Serial.println(fingerID);
  }
  delay(50);
}

int getFingerprintID() {
  uint8_t p = finger.getImage();
  if (p != FINGERPRINT_OK) return -1;

  p = finger.image2Tz();
  if (p != FINGERPRINT_OK) return -1;

  p = finger.fingerFastSearch();
  if (p != FINGERPRINT_OK) return -1;

  Serial.print("Scanned ID #"); Serial.print(finger.fingerID); 
  Serial.print(" (Confidence: "); Serial.print(finger.confidence); Serial.println(")");
  return finger.fingerID;
}