#include <ESP32Servo.h>
#include <Bluepad32.h>
#include <HardwareSerial.h>

HardwareSerial mySerial(2); // UART2 para ESP32

// PIN CONNECTIONS
int ledPin = 2;
int xServoPin = 18; int yServoPin = 19;
int x2ServoPin = 12; int y2ServoPin = 13;
int RAHServoPin = 25; int RACServoPin = 32; int RALServoPin = 22;
int LAHServoPin = 26; int LACServoPin = 33; int LALServoPin = 23;

Servo xServo, yServo, x2Servo, y2Servo;
Servo RAHServo, LAHServo, RACServo, LACServo, RALServo, LALServo;

ControllerPtr myControllers[BP32_MAX_GAMEPADS];

// Inercia individual
float inertiaHead = 0.3;
float inertiaMotors = 0.2;
float inertiaArms = 0.05;
float inertiaCodos = 0.1;

// Suavizado individual
float smoothX = 90, smoothY = 90;                   // Cabeza
float smoothX2 = 90, smoothY2 = 90;                 // Motores
float smoothRAH = 90, smoothLAH = 90;               // Hombros
float smoothRAL = 90, smoothLAL = 90;               // Brazos largos
float smoothRAC = 100, smoothLAC = 40;              // Codos

void onConnectedController(ControllerPtr ctl) {
  for (int i = 0; i < BP32_MAX_GAMEPADS; i++) {
    if (myControllers[i] == nullptr) {
      myControllers[i] = ctl;
      break;
    }
  }
}

void onDisconnectedController(ControllerPtr ctl) {
  for (int i = 0; i < BP32_MAX_GAMEPADS; i++) {
    if (myControllers[i] == ctl) {
      myControllers[i] = nullptr;
      break;
    }
  }
}

void processGamepad(ControllerPtr ctl) {
  digitalWrite(ledPin, ctl->buttons() == 0x0001 ? HIGH : LOW);

 // Cabeza
  //int ejeX = ctl->axisX() * 0.5;    /// problemas interferencia servos o mal mapeo o sw no es el peso ni los pines
  //int ejeY = ctl->axisY() * 0.5;


  // Limitar el rango de movimientos de la cabeza
int ejeX = ctl->axisX();
int ejeY = ctl->axisY();

// Zona muerta para evitar micro-movimientos
if (abs(ejeX) < 20) ejeX = 0;
if (abs(ejeY) < 20) ejeY = 0;

int targetX = map(ejeX, -512, 512, 0, 180);
int targetY = map(ejeY, -512, 512, 0, 180);

// Suavizado exponencial
smoothX += (targetX - smoothX) * inertiaHead;
smoothY += (targetY - smoothY) * inertiaHead;

xServo.write(smoothX);
yServo.write(smoothY);


  // Motores con mezcla
  int ejeRX = ctl->axisRX() * 0.5;
  int ejeRY = ctl->axisRY() * 0.5;
  int canalIzq = constrain(ejeRY + ejeRX, -512, 512);
  int canalDer = constrain(ejeRY - ejeRX, -512, 512);
  int targetX2 = map(canalIzq, -512, 512, 0, 180);
  int targetY2 = map(canalDer, -512, 512, 0, 180);
  smoothX2 = smoothX2 * (1 - inertiaMotors) + targetX2 * inertiaMotors;
  smoothY2 = smoothY2 * (1 - inertiaMotors) + targetY2 * inertiaMotors;
  x2Servo.write(smoothX2);
  y2Servo.write(smoothY2);

  // Hombros (frenado y aceleración)
  int targetLAH = map(ctl->brake(), 0, 1023, 0, 80);
  int targetRAH = map(ctl->throttle(), 0, 1023, 80, 0);
  smoothLAH = smoothLAH * (1 - inertiaArms) + targetLAH * inertiaArms;
  smoothRAH = smoothRAH * (1 - inertiaArms) + targetRAH * inertiaArms;
  LAHServo.write(smoothLAH);
  RAHServo.write(smoothRAH);

  // Botones para codos y brazos largos
  float targetRAC = (ctl->buttons() == 0x0020) ? 40 : 100;
  float targetLAC = (ctl->buttons() == 0x0010) ? 100 : 40;
  float targetRAL = (ctl->buttons() == 0x0002) ? 10 : 100;
  float targetLAL = (ctl->buttons() == 0x0004) ? 180 : 90;

  smoothRAC = smoothRAC * (1 - inertiaCodos) + targetRAC * inertiaCodos;
  smoothLAC = smoothLAC * (1 - inertiaCodos) + targetLAC * inertiaCodos;
  smoothRAL = smoothRAL * (1 - inertiaArms) + targetRAL * inertiaArms;
  smoothLAL = smoothLAL * (1 - inertiaArms) + targetLAL * inertiaArms;

  RACServo.write(smoothRAC);
  LACServo.write(smoothLAC);
  RALServo.write(smoothRAL);
  LALServo.write(smoothLAL);

  // LED secundario
  digitalWrite(ledPin, ctl->buttons() == 0x0002 ? HIGH : LOW);

  // Gestos con dpad
  switch (ctl->dpad()) {
    case 0x02: mySerial.println("n"); break;
    case 0x08: mySerial.println("e"); break;
    case 0x04: mySerial.println("c"); break;
  }
}

void processControllers() {
  for (auto ctl : myControllers) {
    if (ctl && ctl->isConnected() && ctl->hasData() && ctl->isGamepad()) {
      processGamepad(ctl);
    }
  }
}

void setup() {
  pinMode(ledPin, OUTPUT);

  xServo.attach(xServoPin); yServo.attach(yServoPin);
  x2Servo.attach(x2ServoPin); y2Servo.attach(y2ServoPin);
  RAHServo.attach(RAHServoPin); LAHServo.attach(LAHServoPin);
  RACServo.attach(RACServoPin); LACServo.attach(LACServoPin);
  RALServo.attach(RALServoPin); LALServo.attach(LALServoPin);
Serial.begin(9600);
  mySerial.begin(9600, SERIAL_8N1, 16, 17);
  BP32.setup(&onConnectedController, &onDisconnectedController);
  BP32.forgetBluetoothKeys();
  BP32.enableVirtualDevice(false);
}

void loop() {
  if (BP32.update()) processControllers();
  delay(10);
}
