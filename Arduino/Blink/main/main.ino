/*
  Data: 05/01/2025
  Autor: DALÇÓQUIO AUTOMAÇÃO
  Projeto: Supervisório em Visual Basic para Arduino Uno
  Descrição: Recebe valores de variável
 
*/

///////////////////////////////////////////////////////////////////
// FUNÇÃO DE SETUP
void setup() {
  Serial.begin(9600);

  pinMode(13, OUTPUT);

  Serial.println("V0:Setup finalizado.");
  delay(100);
  
}// end setup

///////////////////////////////////////////////////////////////////
// FUNÇÃO DE LOOP
void loop() {
    static bool value_variavel;

    digitalWrite(13, HIGH);
    Serial.println("V1:1");
    delay(100);
    for (int x=0; x<10; x++) {
      Serial.println("V2:" + String(x));
      delay(100);
    }

    digitalWrite(13, LOW);
    Serial.println("V1:0");
    delay(100);
    for (int x=0; x<10; x++) {
      Serial.println("V3:" + String(x));
      delay(100);
    }

    value_variavel = !value_variavel;
    Serial.println("V4:" + String(value_variavel));
    delay(100);

} // end loop
