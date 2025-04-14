#include "LedControl.h"
#include <Servo.h>

Servo servoMotor1;
Servo servoMotor2;

char dato = 0; 
int val = 0;
int val2 = 0;
int inPin = 3;
int variable = 0;

LedControl lc = LedControl(12, 11, 10, 2);  // Pins: DIN, CLK, CS, # of Display connected

unsigned long delayTime = 5;  // Delay between Frames
unsigned long lastCommandTime = 0;
unsigned long blinkInterval = 5000;

// Definir frames de animaciones
byte blink1[] = {
  B00111100,
  B01000010,
  B10000001,
  B10011001,
  B10011001,
  B10000001,
  B01000010,
  B00111100
};

byte blink2[] = {
  B00000000,
  B00111100,
  B01111110,
  B01111110,
  B01111110,
  B01111110,
  B00111100,
  B00000000
};

byte blink3[] = {
  B00000000,
  B00000000,
  B00000000,
  B00111100,
  B01111110,
  B00111100,
  B00000000,
  B00000000
};

// Put values in arrays
byte eye1a[] = {
   B00000000,  
   B00000000,
   B00000000,
   B00011000,
   B00011000,
   B00000000,
   B00000000,
   B00000000
};

byte eye1b[] = {
   B00000000,  
   B00000000,
   B00000000,
   B00011000,
   B00011000,
   B00000000,
   B00000000,
   B00000000
};

// Animaciones adicionales para "corazón", "normal" y "enojado"
byte eye1heart[] = {
  B00001100,
  B00010010,
  B00100010,
  B01000100,
  B01000100,
  B00100010,
  B00010010,
  B00001100
};

byte eye2enojado[] = {
  B00110000,
  B01001000,
  B10000100,
  B10010010,
  B10011001,
  B10000001,
  B01000010,
  B00111100
};

byte eye1enojado[]= //enojado
  {
    B00111100,
    B01000010,
    B10000001,
    B10011001,
    B10010010,
    B10000100,
    B01001000,
    B00110000
  }; 



byte eye1normal[] = {
  B00111100,
  B01000010,
  B10000001,
  B10011001,
  B10011001,
  B10000001,
  B01000010,
  B00111100
};

byte eye2a[] = {
  B00000000,  
  B00000000,
  B00000000,
  B00011000,
  B00011000,
  B00000000,
  B00000000,
  B00000000
};

byte eye1sorpresa[] = {
  B00000000,  // Ojo sorpresa mirando hacia la derecha
  B00000000,
  B00001110,
  B00011111,
  B00011111,
  B00001110,
  B00000000,
  B00000000
};


void setup() {
  Serial.begin(9600);
  servoMotor1.attach(2);
  servoMotor2.attach(3);
  
  lc.shutdown(0, false);  // Wake up displays
  lc.shutdown(1, false);
  lc.setIntensity(0, 5);  // Set intensity levels
  lc.setIntensity(1, 5);
  lc.clearDisplay(0);  // Clear Displays
  lc.clearDisplay(1);
}

void blinkEyes() {
  byte* frames[] = { blink1, blink2, blink3, blink2, blink1};  // Cerrar y abrir
  for (int f = 0; f < 5; f++) {
    for (int i = 0; i < 8; i++) {
      lc.setRow(0, i, frames[f][i]);
      lc.setRow(1, i, frames[f][i]);
    }
    delay(80);
  }
}

// Funciones de animaciones de ojos
void seye1a() {
  for (int i = 0; i < 8; i++) {
    lc.setRow(0, i, eye1a[i]);
  }
}

void seye2a() {
  for (int i = 0; i < 8; i++) {
    lc.setRow(1, i, eye2a[i]);
  }
}

void normal() {
  // Animación normal de ojos
  for (int i = 0; i < 8; i++) {
    lc.setRow(0, i, eye1normal[i]);
    lc.setRow(1, i, eye1normal[i]);
  }
  servoMotor1.write(40);
  servoMotor2.write(150);
}

void corazon() {
  // Animación de corazón en los ojos
  for (int i = 0; i < 8; i++) {
    lc.setRow(0, i, eye1heart[i]);
    lc.setRow(1, i, eye1heart[i]);
  }
  servoMotor1.write(60);
  servoMotor2.write(150);
}

void enojado() {
  // Ojos enojados
  for (int i = 0; i < 8; i++) {
    lc.setRow(0, i, eye1enojado[i]);
    lc.setRow(1, i, eye2enojado[i]);
  }
  servoMotor1.write(60);
  servoMotor2.write(130);
}

void sorpresa() {
  // Ojos sorprendidos
  for (int i = 0; i < 8; i++) {
    lc.setRow(0, i, eye1sorpresa[i]);
    lc.setRow(1, i, eye1sorpresa[i]);
  }
  servoMotor1.write(50);
  servoMotor2.write(150);
}

void loop() {
  if (Serial.available()) {
    char dato = Serial.read();
    lastCommandTime = millis();  // Reinicia el contador de inactividad

    if (dato == 'n') {
      normal();
    } else if (dato == 'c') {
      corazon();
    } else if (dato == 'e') {
      enojado();
    
     } else if (dato == 's') {
      sorpresa();
    
    }
    
  }

  // Si no hay comandos por más de 5 segundos
  if (millis() - lastCommandTime > blinkInterval) {
    blinkEyes();  // Parpadeo de los ojos
    //normal();     // Vuelve a la posición normal después del parpadeo
    lastCommandTime = millis(); // Reinicia después del parpadeo
  }
}
