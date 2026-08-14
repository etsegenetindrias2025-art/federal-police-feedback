#include <WiFi.h>
#include <HTTPClient.h>

const char* ssid = "Wokwi-GUEST";
const char* password = "";

void setup() {
  Serial.begin(115200);
  WiFi.begin(ssid, password);
  
  while (WiFi.status() != WL_CONNECTED) {
    delay(500);
    Serial.print(".");
  }
  Serial.println("Connected to Virtual Wi-Fi!");
}

void loop() {
  if (WiFi.status() == WL_CONNECTED) {
    HTTPClient http;
    // Point to your local Flask app running in VS Code
    http.begin("http://127.0.0.1:5000/api/sensor-data");
    http.addHeader("Content-Type", "application/json");
    
    int httpResponseCode = http.POST("{\"presence\": 1}");
    
    if (httpResponseCode > 0) {
      String response = http.getString();
      Serial.println(httpResponseCode);
      Serial.println(response);
    }
    http.end();
  }
  delay(10000); // Send data every 10 seconds
}