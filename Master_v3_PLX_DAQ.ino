// MASTER v3 - MODIFICADO PARA PLX-DAQ v2
// Envia 12 temperaturas (6 locales + 6 remotas) a Excel
#include <Wire.h>
#include <max6675.h>

#define SLAVE_ADDR 8

int pinSO  = 4;
int pinSCK = 6;

const int numLocal = 6;
int CS_local[numLocal] = {22,24,26,28,30,32};

MAX6675* sensores[numLocal];
float tempLocal[6];
float tempSlave[6];

void setup() {
  Serial.begin(9600);
  Wire.begin();

  for(int i=0; i<numLocal; i++){
    sensores[i] = new MAX6675(pinSCK, CS_local[i], pinSO);
  }

  delay(500);

  // PLX-DAQ: limpiar y definir columnas
  Serial.println("CLEARDATA");
  Serial.println("LABEL,Timestamp,Local_1,Local_2,Local_3,Local_4,Local_5,Local_6,Remoto_1,Remoto_2,Remoto_3,Remoto_4,Remoto_5,Remoto_6");
}

void loop() {
  // Leer sensores locales
  for(int i=0; i<6; i++){
    tempLocal[i] = sensores[i]->readCelsius();
  }

  // Leer sensores remotos via I2C
  Wire.requestFrom(SLAVE_ADDR, 32);
  for(int i=0; i<6; i++){
    byte* p = (byte*)&tempSlave[i];
    for(int j=0; j<4; j++){
      if(Wire.available()){
        p[j] = Wire.read();
      }
    }
  }

  // Enviar datos en formato PLX-DAQ
  Serial.print("DATA,DATE,TIME,");
  for(int i=0; i<6; i++){
    Serial.print(tempLocal[i]);
    Serial.print(",");
  }
  for(int i=0; i<5; i++){
    Serial.print(tempSlave[i]);
    Serial.print(",");
  }
  Serial.println(tempSlave[5]);

  delay(2000);
}
