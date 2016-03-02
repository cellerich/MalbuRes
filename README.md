# MalbuRes
Flugschul Reservationssystem mit JotForm, Google Spreadsheet und Scripts


Mit einem JotForm Formular welches seine Daten an ein Google Spreadsheet sendet wird ein Reservationsprozess ausgelöst
Der Prozess läuft als "onchange" Code im Google Spreadsheet und hat folgende Funtionen:
- Überprüfen ob im besagtem Zeitraum bereits eine Buchung existiert
- Wenn nicht:
-- Kalendereintrag in Google Kalender vornehmen 
-- Bestätigungsmail an Absender senden
-- ggf. Bestätigungsmail an Fluglehrer senden
- wenn ja:
-- Reservation per Mail zurückweisen

Ein weiterer Prozess (Reservation löschen) kann durch einen Link im Bestätigungsmail angeworfen werden:
- Kalender Eintrag löschen
- Bestätigung der Löschung per Mail an Absender
- ggf. Absagemail an Fluglehrer senden


