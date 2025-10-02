# TIA Openness Exporter

Jednoduchý nástroj v C# pro export PLC bloků, tagů a konfigurace z **Siemens TIA Portal (V19)** pomocí **Openness API**.

## Funkce
- Export všech FB/FC/DB bloků do XML
- Export symbolických tagů
- Připraveno na rozšíření (HW topology, HMI atd.)

## Použití
1. Otevři projekt ve Visual Studiu (Console App .NET Framework 4.8).
2. Přidej reference na knihovny z TIA PublicAPI:
   - `Siemens.Engineering.dll`
   - `Siemens.Engineering.Hmi.dll`
3. Zkompiluj a spusť:
   ```bash
   TIAExporter.exe "C:\Path\Project.ap19" "C:\Exports"
