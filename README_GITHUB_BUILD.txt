ALKIS_GeoAS_GEOgraf V1.2.0 – GitHub-Bauanleitung
=================================================

VORHER
------
In einem bereits verwendeten Repository unter .github/workflows alle alten
Workflow-Dateien löschen. Danach darf dort nur Build-windows-exe.yml liegen.

HOCHLADEN
---------
1. Diese ZIP-Datei vollständig entpacken.
2. Im GitHub-Repository Add file -> Upload files öffnen.
3. Den gesamten entpackten Inhalt hochladen.
4. Darauf achten, dass der Ordner .github mit hochgeladen wird.
5. Commit changes bestätigen.

EXE ERZEUGEN
------------
1. Den Reiter Actions öffnen.
2. Links Windows-EXE erstellen auswählen.
3. Run workflow und danach erneut Run workflow drücken.
4. Nach dem grünen Abschluss das Artefakt
   ALKIS_GeoAS_GEOgraf_V1.2.0_Windows herunterladen.

ERGEBNIS
--------
Das heruntergeladene GitHub-Artefakt ist bereits das einzige ZIP.
Nach einmaligem Entpacken liegen die EXE, SHA256.txt, BUILD_INFO.txt und die
Begleitunterlagen vor. Ein zweites ZIP wird nicht erzeugt.

SICHERHEITSPRÜFUNG
------------------
Vor dem EXE-Bau führt GitHub verify_v120.py aus. Dieser Test bricht den Build ab,
wenn der Quellcode nicht Version 1.2.0 ist oder HTML-Starter, Dezimalanteile,
1302/1102, gemeinsame Buchungskomplexe beziehungsweise die Integritätsprüfung
nicht funktionieren.
