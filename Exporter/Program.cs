using Siemens.Engineering;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW.Tags;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Exporter
{
    internal class Program
    {
        // Volitelný blacklist názvů bloků (prefixy). Pokud něco vadí, přidej sem.
        private static readonly string[] BlockNameBlacklistPrefixes = new string[0]; // C# 7.3: explicitní typ

        static int Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Použití: Exporter.exe <Project.ap19> <ExportDir>");
                return 1;
            }

            string projectPath = args[0];
            string exportDir = args[1];

            try
            {
                if (!File.Exists(projectPath))
                {
                    Console.WriteLine("Chyba: Soubor projektu neexistuje: " + projectPath);
                    return 2;
                }

                Directory.CreateDirectory(exportDir);

                // Log do souboru
                string logPath = Path.Combine(exportDir, "export_log.txt");
                using (var logFs = new FileStream(logPath, FileMode.Create, FileAccess.Write, FileShare.Read))
                using (var logSw = new StreamWriter(logFs, new UTF8Encoding(false)) { AutoFlush = true })
                {
                    Action<string> LOG = s => { Console.WriteLine(s); logSw.WriteLine(s); };

                    LOG("== Start ==");
                    LOG("Project: " + projectPath);
                    LOG("ExportDir: " + exportDir);

                    // CSV pro bloky, které nejdou exportovat (např. Safety)
                    string notExportedCsv = Path.Combine(exportDir, "not_exported.csv");
                    if (!File.Exists(notExportedCsv))
                        File.WriteAllText(notExportedCsv, "Device,BlockName,Kind,Reason\n", new UTF8Encoding(false));

                    using (var portal = new TiaPortal(TiaPortalMode.WithoutUserInterface))
                    {
                        LOG("Otevírám projekt…");
                        var project = portal.Projects.Open(new FileInfo(projectPath));
                        LOG("Projekt otevřen.");

                        int blocksTotal = 0;
                        int tagsTotal = 0;

                        LOG("Zařízení v projektu: " + project.Devices.Count);

                        // 3) Export HW topologie (CSV pro celé project)
                        ExportHardwareCsv(project, exportDir, LOG);

                        foreach (var device in project.Devices)
                        {
                            LOG("- Device: " + device.Name + ", items: " + device.DeviceItems.Count);

                            foreach (var di in device.DeviceItems)
                            {
                                var sw = di.GetService<SoftwareContainer>()?.Software;
                                if (sw is PlcSoftware plc)
                                {
                                    LOG("  • PLC software nalezen v '" + device.Name + "' → " + plc.Name);

                                    // podsložka pro device
                                    string devDir = Path.Combine(exportDir, MakeSafeFileName(device.Name));
                                    Directory.CreateDirectory(devDir);

                                    // 1) Export bloků do XML (rekurzivně)
                                    int exportedBlocks = ExportBlocksRecursively(plc.BlockGroup, devDir, device.Name, notExportedCsv, LOG);
                                    blocksTotal += exportedBlocks;
                                    LOG("    Export bloků: " + exportedBlocks + " ks → " + devDir);

                                    // 2) Export tagů do CSV (append do jednoho souboru v rootu exportu)
                                    string csvPath = Path.Combine(exportDir, "plc_tags.csv");
                                    bool writeHeader = !File.Exists(csvPath);

                                    using (var fs = new FileStream(csvPath, FileMode.Append, FileAccess.Write, FileShare.Read))
                                    using (var swCsv = new StreamWriter(fs, new UTF8Encoding(false)))
                                    {
                                        if (writeHeader)
                                            swCsv.WriteLine("Name,Address,DataType,Comment,Table,GroupPath,Device");

                                        int exportedTags = ExportTagTablesCsv(plc.TagTableGroup, swCsv, device.Name, LOG);
                                        tagsTotal += exportedTags;
                                        LOG("    Export tagů: " + exportedTags + " ks → " + csvPath);
                                    }
                                }
                                else
                                {
                                    LOG("  • (skip) V '" + device.Name + "' nebyl PLC software (HMI/komponenta).");
                                }
                            }
                        }

                        project.Close();
                        LOG("== Hotovo: bloků " + blocksTotal + ", tagů " + tagsTotal + " ==");
                    }

                    // indikátor výstupu
                    File.AppendAllText(Path.Combine(exportDir, ".ok"), DateTime.Now.ToString("s"));
                }

                Console.WriteLine("Hotovo → " + exportDir);
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Chyba: " + ex.Message);
                return 99;
            }
        }

        // =========================
        // EXPORT BLOKŮ (XML)
        // =========================

        private static int ExportBlocksRecursively(PlcBlockSystemGroup group, string exportDir, string deviceName, string notExportedCsv, Action<string> LOG)
        {
            int count = 0;

            LOG("    [Group:System] Blocks=" + group.Blocks.Count + ", SubGroups=" + group.Groups.Count);
            foreach (PlcBlock block in group.Blocks)
            {
                if (ShouldSkip(block))
                {
                    File.AppendAllText(notExportedCsv, $"{SafeCsv(deviceName)},{SafeCsv(block.Name)},{SafeCsv(BlockKind(block))},SkippedByRule\n");
                    continue;
                }

                if (ExportOneBlock(block, exportDir, deviceName, notExportedCsv, LOG)) count++;
            }

            foreach (PlcBlockUserGroup sub in group.Groups)
            {
                count += ExportBlocksRecursively(sub, exportDir, deviceName, notExportedCsv, LOG);
            }
            return count;
        }

        private static int ExportBlocksRecursively(PlcBlockUserGroup group, string exportDir, string deviceName, string notExportedCsv, Action<string> LOG)
        {
            int count = 0;

            LOG("    [Group:User:" + group.Name + "] Blocks=" + group.Blocks.Count + ", SubGroups=" + group.Groups.Count);
            foreach (PlcBlock block in group.Blocks)
            {
                if (ShouldSkip(block))
                {
                    File.AppendAllText(notExportedCsv, $"{SafeCsv(deviceName)},{SafeCsv(block.Name)},{SafeCsv(BlockKind(block))},SkippedByRule\n");
                    continue;
                }

                if (ExportOneBlock(block, exportDir, deviceName, notExportedCsv, LOG)) count++;
            }

            foreach (PlcBlockUserGroup sub in group.Groups)
            {
                count += ExportBlocksRecursively(sub, exportDir, deviceName, notExportedCsv, LOG);
            }
            return count;
        }

        private static bool ExportOneBlock(PlcBlock block, string exportDir, string deviceName, string notExportedCsv, Action<string> LOG)
        {
            try
            {
                string safeName = MakeSafeFileName(BuildBlockFileName(block));
                string filePath = Path.Combine(exportDir, safeName + ".xml");
                LOG("      Exportuji blok: " + block.Name + " → " + filePath);
                block.Export(new FileInfo(filePath), ExportOptions.WithDefaults);
                return true;
            }
            catch (Exception ex)
            {
                LOG("      ! Nepodařilo se exportovat '" + block.Name + "': " + ex.Message);
                // .NET Framework: použijeme IndexOf místo Contains(String,StringComparison)
                bool notPerm = ex.Message != null && ex.Message.IndexOf("not permitted", StringComparison.OrdinalIgnoreCase) >= 0;
                var reason = notPerm ? "NotPermitted" : "Error";
                File.AppendAllText(notExportedCsv, $"{SafeCsv(deviceName)},{SafeCsv(block.Name)},{SafeCsv(BlockKind(block))},{SafeCsv(reason + ": " + ex.Message)}\n");
                return false;
            }
        }

        private static bool ShouldSkip(PlcBlock block)
        {
            foreach (var p in BlockNameBlacklistPrefixes)
                if (!string.IsNullOrEmpty(p) && block.Name.StartsWith(p, StringComparison.OrdinalIgnoreCase))
                    return true;

            return false;
        }

        private static string BlockKind(PlcBlock block)
        {
            string kind = block.GetType().Name; // PlcBlockFb/Fc/Db/Ob...
            if (kind.Contains("Fb")) return "FB";
            if (kind.Contains("Fc")) return "FC";
            if (kind.Contains("Db")) return "DB";
            if (kind.Contains("Ob")) return "OB";
            return kind;
        }

        private static string BuildBlockFileName(PlcBlock block)
        {
            string kind = BlockKind(block);
            string number = FirstNonEmpty(TryGetAttr(block, "Number"), TryGetProp(block, "Number"));
            return string.IsNullOrWhiteSpace(number)
                ? (kind + "_" + block.Name)
                : (kind + "_" + number + "_" + block.Name);
        }

        // =========================
        // EXPORT TAGŮ (CSV)
        // =========================

        private static int ExportTagTablesCsv(PlcTagTableSystemGroup root, StreamWriter sw, string deviceName, Action<string> LOG)
        {
            int count = 0;

            LOG("    [Tags:SystemGroup] TagTables=" + root.TagTables.Count + ", Groups=" + root.Groups.Count);
            foreach (PlcTagTable table in root.TagTables)
                count += WriteTableTags(sw, table, "", deviceName, LOG);

            foreach (PlcTagTableUserGroup g in root.Groups)
                count += ExportTagTablesCsv(g, sw, deviceName, g.Name, LOG);

            return count;
        }

        private static int ExportTagTablesCsv(PlcTagTableUserGroup group, StreamWriter sw, string deviceName, string groupPath, Action<string> LOG)
        {
            int count = 0;

            LOG("    [Tags:UserGroup:" + group.Name + "] TagTables=" + group.TagTables.Count + ", Groups=" + group.Groups.Count);
            foreach (PlcTagTable table in group.TagTables)
                count += WriteTableTags(sw, table, groupPath, deviceName, LOG);

            foreach (PlcTagTableUserGroup sub in group.Groups)
            {
                string subPath = string.IsNullOrEmpty(groupPath) ? sub.Name : (groupPath + "/" + sub.Name);
                count += ExportTagTablesCsv(sub, sw, deviceName, subPath, LOG);
            }
            return count;
        }

        private static int WriteTableTags(StreamWriter sw, PlcTagTable table, string groupPath, string deviceName, Action<string> LOG)
        {
            int count = 0;
            try
            {
                LOG("      TagTable: " + table.Name + ", Tags=" + table.Tags.Count);
                foreach (PlcTag tag in table.Tags)
                {
                    string name = FirstNonEmpty(TryGetProp(tag, "Name"), TryGetAttr(tag, "Name"));
                    string address = FirstNonEmpty(TryGetProp(tag, "LogicalAddress"), TryGetAttr(tag, "LogicalAddress"),
                                                    TryGetProp(tag, "Address"), TryGetAttr(tag, "Address"));
                    string dataType = FirstNonEmpty(TryGetProp(tag, "DataTypeName"), TryGetAttr(tag, "DataTypeName"),
                                                    TryGetProp(tag, "DataType"), TryGetAttr(tag, "DataType"));
                    string comment = FirstNonEmpty(TryGetProp(tag, "Comment"), TryGetAttr(tag, "Comment"));

                    sw.WriteLine(string.Join(",",
                        SafeCsv(name),
                        SafeCsv(address),
                        SafeCsv(dataType),
                        SafeCsv(comment),
                        SafeCsv(table.Name),
                        SafeCsv(groupPath ?? ""),
                        SafeCsv(deviceName)
                    ));
                    count++;
                }
            }
            catch (Exception ex)
            {
                LOG("      ! TagTable '" + table.Name + "' přeskočena: " + ex.Message);
            }
            return count;
        }

        // =========================
        // EXPORT HW TOPOLOGIE (CSV)
        // =========================

        private static void ExportHardwareCsv(Project project, string exportDir, Action<string> LOG)
        {
            string hwCsv = Path.Combine(exportDir, "hardware.csv");
            bool writeHeader = !File.Exists(hwCsv);

            using (var fs = new FileStream(hwCsv, FileMode.Append, FileAccess.Write, FileShare.Read))
            using (var sw = new StreamWriter(fs, new UTF8Encoding(false)))
            {
                if (writeHeader)
                    sw.WriteLine("Device,ItemPath,ItemType,OrderNumber,Rack,Slot,PN_Name,IP,Subnet,Comment");

                foreach (var device in project.Devices)
                {
                    ExportDeviceItemsRecursive(device, device.DeviceItems, device.Name, "", sw, LOG);
                }
            }

            LOG("HW topologie → " + hwCsv);
        }

        private static void ExportDeviceItemsRecursive(Device device, DeviceItemComposition items, string deviceName, string parentPath, StreamWriter sw, Action<string> LOG)
        {
            foreach (var item in items)
            {
                string itemName = FirstNonEmpty(TryGetProp(item, "Name"), TryGetAttr(item, "Name"));
                string typeName = item.GetType().Name;
                string itemPath = string.IsNullOrEmpty(parentPath) ? itemName : (parentPath + "/" + itemName);

                string orderNo = FirstNonEmpty(TryGetAttr(item, "OrderNumber"), TryGetProp(item, "OrderNumber"));
                string rack = FirstNonEmpty(TryGetAttr(item, "RackNumber"), TryGetProp(item, "RackNumber"));
                string slot = FirstNonEmpty(TryGetAttr(item, "SlotNumber"), TryGetProp(item, "SlotNumber"));
                string comment = FirstNonEmpty(TryGetAttr(item, "Comment"), TryGetProp(item, "Comment"));

                string pnName = FirstNonEmpty(TryGetAttr(item, "ProfinetDeviceName"), TryGetProp(item, "ProfinetDeviceName"),
                                                 TryGetAttr(item, "PNName"), TryGetProp(item, "PNName"));
                string ip = FirstNonEmpty(TryGetAttr(item, "IPV4Address"), TryGetProp(item, "IPV4Address"),
                                                 TryGetAttr(item, "IPv4Address"), TryGetProp(item, "IPv4Address"),
                                                 TryGetAttr(item, "IpAddress"), TryGetProp(item, "IpAddress"));
                string subnet = FirstNonEmpty(TryGetAttr(item, "SubnetMask"), TryGetProp(item, "SubnetMask"));

                sw.WriteLine(string.Join(",",
                    SafeCsv(deviceName),
                    SafeCsv(itemPath),
                    SafeCsv(typeName),
                    SafeCsv(orderNo),
                    SafeCsv(rack),
                    SafeCsv(slot),
                    SafeCsv(pnName),
                    SafeCsv(ip),
                    SafeCsv(subnet),
                    SafeCsv(comment)
                ));

                ExportDeviceItemsRecursive(device, item.DeviceItems, deviceName, itemPath, sw, LOG);
            }
        }

        // =========================
        // Utility (reflexe + CSV + názvy)
        // =========================

        private static string TryGetProp(object obj, string propName)
        {
            try
            {
                var pi = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
                if (pi == null) return "";
                var val = pi.GetValue(obj, null);
                return Convert.ToString(val) ?? "";
            }
            catch { return ""; }
        }

        private static string TryGetAttr(object obj, string attrName)
        {
            try
            {
                var mi = obj.GetType().GetMethod("GetAttribute",
                    BindingFlags.Public | BindingFlags.Instance,
                    null, new[] { typeof(string) }, null);
                if (mi == null) return "";
                var val = mi.Invoke(obj, new object[] { attrName });
                return Convert.ToString(val) ?? "";
            }
            catch { return ""; }
        }

        private static string FirstNonEmpty(params string[] values)
        {
            foreach (var v in values)
                if (!string.IsNullOrWhiteSpace(v)) return v;
            return "";
        }

        private static string MakeSafeFileName(string name)
        {
            var invalid = Path.GetInvalidFileNameChars();
            var chars = name.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
                if (Array.IndexOf(invalid, chars[i]) >= 0) chars[i] = '_';
            var safe = new string(chars);
            return safe.Length > 200 ? safe.Substring(0, 200) : safe;
        }

        private static string SafeCsv(string s)
        {
            if (s == null) return "";
            bool needsQuotes = s.IndexOfAny(new[] { ',', '"', '\n', '\r' }) >= 0;
            s = s.Replace("\"", "\"\"");
            return needsQuotes ? "\"" + s + "\"" : s;
        }
    }
}
