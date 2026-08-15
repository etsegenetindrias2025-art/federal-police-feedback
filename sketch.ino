#include <WiFi.h>
#include <HTTPClient.h>

// Wi-Fi credentials (configured for Wokwi simulation)
const char* ssid = "Wokwi-GUEST";
const char* password = "";

// Server endpoint 
// Note: If running in Wokwi simulator, use "http://host.wokwi.internal:5000/api/iot/sensor-event" 
// or your computer's local network IP address (e.g., "http://192.168.1.100:5000/api/iot/sensor-event")
const char* serverUrl = "http://host.wokwi.internal:5000/api/iot/sensor-event";

// Pin definitions
const int pirPin = 14;       // PIR motion sensor pin
const int buzzerPin = 27;    // Buzzer pin
const int ledRed = 26;       // Red LED pin
const int ledGreen = 25;     // Green LED pin

void setup() {
  Serial.begin(115200);
  
  pinMode(pirPin, INPUT);
  pinMode(buzzerPin, OUTPUT);
  pinMode(ledRed, OUTPUT);
  pinMode(ledGreen, OUTPUT);

  // Connect to Wi-Fi
  WiFi.begin(ssid, password);
  Serial.print("Connecting to WiFi");
  while (WiFi.status() != WL_CONNECTED) {
    delay(500);
    Serial.print(".");
  }
  Serial.println("\nWiFi connected.");
}

void loop() {
  int pirState = digitalRead(pirPin);

  if (pirState == HIGH) {
    Serial.println("Motion detected by PIR!");
    
    // Trigger local feedback indicator
    digitalWrite(ledGreen, HIGH);
    digitalWrite(buzzerPin, HIGH);
    delay(200);
    digitalWrite(buzzerPin, LOW);
    
    // Send HTTP POST request to Flask backend
    if (WiFi.status() == WL_CONNECTED) {
      HTTPClient http;
      http.begin(serverUrl);
      http.addHeader("Content-Type", "application/json");
      
      String payload = "{\"sensor_type\": \"pir\", \"device_id\": \"esp32_kiosk_01\"}";
      int httpResponseCode = http.POST(payload);
      
      if (httpResponseCode > 0) {
        String response = http.getString();
        Serial.println("Server Response Code: " + String(httpResponseCode));
        Serial.println("Server Response: " + response);
      } else {
        Serial.print("Error on sending POST: ");
        Serial.println(httpResponseCode);
      }
      http.end();
    }
    
    digitalWrite(ledGreen, LOW);
    delay(5000); // Cooldown to prevent duplicate triggers
  }
  
  delay(100);
}