using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using ClosedXML.Excel;

namespace RebarCadSync
{
    public class RebarSync
    {
        private const double DefaultRadius = 350.0;
        private static readonly Regex SpecRegex = new Regex(@"^\s*\d+\s*HA\s*\d+\s*$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex LengthRegex = new Regex(@"^\s*\d+(\.\d+)?\s*$", RegexOptions.Compiled);

        [CommandMethod("REBARSYNC")]
        public void Run()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                return;
            }

            var editor = doc.Editor;
            editor.WriteMessage("\nREBARSYNC: Please save a DWG backup before running this command.");

            var excelPath = PromptForExcelPath(editor);
            if (string.IsNullOrWhiteSpace(excelPath))
            {
                return;
            }

            if (!File.Exists(excelPath))
            {
                editor.WriteMessage($"\nExcel file not found: {excelPath}");
                return;
            }

            Dictionary<string, RebarRow> excelRows;
            try
            {
                excelRows = LoadExcel(excelPath);
            }
            catch (Exception ex)
            {
                editor.WriteMessage($"\nFailed to read Excel: {ex.Message}");
                return;
            }

            if (excelRows.Count == 0)
            {
                editor.WriteMessage("\nNo rebar rows found in Excel sheet 'Rebar'.");
                return;
            }

            var db = doc.Database;
            var changes = new List<ChangeRecord>();
            var ambiguousMarks = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var missingMarks = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var updatedMarks = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            using (var transaction = db.TransactionManager.StartTransaction())
            {
                var modelSpace = (BlockTableRecord)transaction.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);
                var texts = LoadTexts(modelSpace, transaction);

                var markTexts = texts
                    .Where(t => t.IsDbText && string.Equals(t.Layer, "Dim", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                foreach (var markText in markTexts)
                {
                    var markKey = markText.Text.Trim();
                    if (!excelRows.TryGetValue(markKey, out var row))
                    {
                        continue;
                    }

                    var specCandidates = FindCandidates(texts, markText, DefaultRadius, SpecRegex, allowSpecMatch: true);
                    var lengthCandidates = FindCandidates(texts, markText, DefaultRadius, LengthRegex, allowSpecMatch: false);

                    if (specCandidates.IsAmbiguous || lengthCandidates.IsAmbiguous)
                    {
                        ambiguousMarks.Add(markKey);
                        changes.Add(ChangeRecord.Ambiguous(markKey, specCandidates, lengthCandidates));
                        continue;
                    }

                    if (specCandidates.Nearest == null || lengthCandidates.Nearest == null)
                    {
                        var missingParts = new List<string>();
                        if (specCandidates.Nearest == null)
                        {
                            missingParts.Add("spec");
                        }

                        if (lengthCandidates.Nearest == null)
                        {
                            missingParts.Add("length");
                        }

                        changes.Add(ChangeRecord.Missing(markKey, string.Join("+", missingParts)));
                        continue;
                    }

                    var specTarget = specCandidates.Nearest;
                    var lengthTarget = lengthCandidates.Nearest;

                    var newSpec = $"{row.Qty}HA{row.Dia}";
                    var newLength = FormatLength(row.LengthCm);

                    var oldSpec = specTarget.Text;
                    var oldLength = lengthTarget.Text;

                    UpdateTextEntity(transaction, specTarget, newSpec);
                    UpdateTextEntity(transaction, lengthTarget, newLength);

                    changes.Add(ChangeRecord.Updated(markKey, oldSpec, newSpec, oldLength, newLength));
                    updatedMarks.Add(markKey);
                }

                transaction.Commit();
            }

            foreach (var mark in excelRows.Keys)
            {
                if (!updatedMarks.Contains(mark) && !ambiguousMarks.Contains(mark))
                {
                    missingMarks.Add(mark);
                }
            }

            var logPath = WriteLog(excelPath, changes, excelRows.Count, updatedMarks.Count, missingMarks, ambiguousMarks);
            editor.WriteMessage($"\nREBARSYNC complete. Log written to: {logPath}");
        }

        private static string PromptForExcelPath(Editor editor)
        {
            var options = new PromptOpenFileOptions("Select Excel file (.xlsx)")
            {
                Filter = "Excel (*.xlsx)|*.xlsx",
                DialogCaption = "Select Rebar Excel File"
            };

            var result = editor.GetFileNameForOpen(options);
            if (result.Status != PromptStatus.OK)
            {
                return null;
            }

            return result.StringResult;
        }

        private static Dictionary<string, RebarRow> LoadExcel(string path)
        {
            var rows = new Dictionary<string, RebarRow>(StringComparer.OrdinalIgnoreCase);
            using (var workbook = new XLWorkbook(path))
            {
                var sheet = workbook.Worksheet("Rebar");
                if (sheet == null)
                {
                    return rows;
                }

                var headerRow = sheet.FirstRowUsed();
                if (headerRow == null)
                {
                    return rows;
                }

                var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (var cell in headerRow.CellsUsed())
                {
                    var header = cell.GetString().Trim();
                    if (!string.IsNullOrWhiteSpace(header))
                    {
                        headers[header] = cell.Address.ColumnNumber;
                    }
                }

                if (!headers.TryGetValue("mark", out var markCol) ||
                    !headers.TryGetValue("qty", out var qtyCol) ||
                    !headers.TryGetValue("dia", out var diaCol) ||
                    !headers.TryGetValue("length_cm", out var lengthCol))
                {
                    throw new InvalidOperationException("Excel sheet 'Rebar' must include columns: mark, qty, dia, length_cm.");
                }

                foreach (var row in sheet.RowsUsed().Skip(1))
                {
                    var mark = row.Cell(markCol).GetString().Trim();
                    if (string.IsNullOrWhiteSpace(mark))
                    {
                        continue;
                    }

                    if (!TryGetInt(row.Cell(qtyCol), out var qty) || !TryGetInt(row.Cell(diaCol), out var dia))
                    {
                        continue;
                    }

                    if (!TryGetDouble(row.Cell(lengthCol), out var length))
                    {
                        continue;
                    }

                    rows[mark] = new RebarRow(mark, qty, dia, length);
                }
            }

            return rows;
        }

        private static bool TryGetInt(IXLCell cell, out int value)
        {
            value = 0;
            if (cell.TryGetValue(out double numeric))
            {
                value = Convert.ToInt32(Math.Round(numeric));
                return true;
            }

            var text = cell.GetString().Trim();
            return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
        }

        private static bool TryGetDouble(IXLCell cell, out double value)
        {
            value = 0;
            if (cell.TryGetValue(out double numeric))
            {
                value = numeric;
                return true;
            }

            var text = cell.GetString().Trim();
            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        private static List<TextInfo> LoadTexts(BlockTableRecord modelSpace, Transaction transaction)
        {
            var list = new List<TextInfo>();
            foreach (ObjectId id in modelSpace)
            {
                var entity = transaction.GetObject(id, OpenMode.ForRead) as Entity;
                if (entity == null)
                {
                    continue;
                }

                if (entity is DBText dbText)
                {
                    list.Add(TextInfo.FromDbText(dbText));
                }
                else if (entity is MText mText)
                {
                    list.Add(TextInfo.FromMText(mText));
                }
            }

            return list;
        }

        private static CandidateResult FindCandidates(List<TextInfo> texts, TextInfo markText, double radius, Regex regex, bool allowSpecMatch)
        {
            var radiusSquared = radius * radius;
            var candidates = new List<TextInfo>();

            foreach (var text in texts)
            {
                if (text.ObjectId == markText.ObjectId)
                {
                    continue;
                }

                var raw = text.Text;
                if (!regex.IsMatch(raw))
                {
                    continue;
                }

                if (!allowSpecMatch && SpecRegex.IsMatch(raw))
                {
                    continue;
                }

                var distSquared = DistanceSquared(markText.Position, text.Position);
                if (distSquared <= radiusSquared)
                {
                    candidates.Add(text);
                }
            }

            if (candidates.Count == 0)
            {
                return CandidateResult.Empty();
            }

            var ordered = candidates.OrderBy(c => DistanceSquared(markText.Position, c.Position)).ToList();
            var nearest = ordered[0];
            var nearestDistance = DistanceSquared(markText.Position, nearest.Position);
            var tieCount = ordered.Count(c => Math.Abs(DistanceSquared(markText.Position, c.Position) - nearestDistance) < 1e-6);

            if (tieCount > 1)
            {
                return CandidateResult.Ambiguous(ordered);
            }

            return CandidateResult.Single(nearest);
        }

        private static void UpdateTextEntity(Transaction transaction, TextInfo info, string newText)
        {
            var entity = (Entity)transaction.GetObject(info.ObjectId, OpenMode.ForWrite);
            if (entity is DBText dbText)
            {
                dbText.TextString = newText;
            }
            else if (entity is MText mText)
            {
                mText.Contents = newText;
            }
        }

        private static string FormatLength(double length)
        {
            var rounded = Math.Round(length);
            if (Math.Abs(length - rounded) < 1e-6)
            {
                return rounded.ToString("0", CultureInfo.InvariantCulture);
            }

            return length.ToString("0.0", CultureInfo.InvariantCulture);
        }

        private static double DistanceSquared(Point3d a, Point3d b)
        {
            var dx = a.X - b.X;
            var dy = a.Y - b.Y;
            var dz = a.Z - b.Z;
            return dx * dx + dy * dy + dz * dz;
        }

        private static string WriteLog(
            string excelPath,
            List<ChangeRecord> changes,
            int excelCount,
            int updatedCount,
            HashSet<string> missingMarks,
            HashSet<string> ambiguousMarks)
        {
            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var fileName = $"RebarSyncLog_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
            var path = Path.Combine(desktop, fileName);

            using (var writer = new StreamWriter(path))
            {
                writer.WriteLine("Section,Detail");
                writer.WriteLine($"Excel File,{excelPath}");
                writer.WriteLine($"Excel Rows,{excelCount}");
                writer.WriteLine($"Updated Marks,{updatedCount}");
                writer.WriteLine($"Missing Marks,{string.Join("|", missingMarks)}");
                writer.WriteLine($"Ambiguous Marks,{string.Join("|", ambiguousMarks)}");
                writer.WriteLine();
                writer.WriteLine("Mark,Status,OldSpec,NewSpec,OldLength,NewLength,Note");

                foreach (var change in changes)
                {
                    writer.WriteLine(change.ToCsv());
                }
            }

            return path;
        }

        private sealed class RebarRow
        {
            public RebarRow(string mark, int qty, int dia, double lengthCm)
            {
                Mark = mark;
                Qty = qty;
                Dia = dia;
                LengthCm = lengthCm;
            }

            public string Mark { get; }
            public int Qty { get; }
            public int Dia { get; }
            public double LengthCm { get; }
        }

        private sealed class TextInfo
        {
            private TextInfo(ObjectId objectId, string text, Point3d position, string layer, bool isDbText)
            {
                ObjectId = objectId;
                Text = text;
                Position = position;
                Layer = layer;
                IsDbText = isDbText;
            }

            public ObjectId ObjectId { get; }
            public string Text { get; }
            public Point3d Position { get; }
            public string Layer { get; }
            public bool IsDbText { get; }

            public static TextInfo FromDbText(DBText dbText)
            {
                return new TextInfo(dbText.ObjectId, dbText.TextString, dbText.Position, dbText.Layer, true);
            }

            public static TextInfo FromMText(MText mText)
            {
                return new TextInfo(mText.ObjectId, mText.Text, mText.Location, mText.Layer, false);
            }
        }

        private sealed class CandidateResult
        {
            private CandidateResult(TextInfo nearest, bool isAmbiguous, List<TextInfo> candidates)
            {
                Nearest = nearest;
                IsAmbiguous = isAmbiguous;
                Candidates = candidates;
            }

            public TextInfo Nearest { get; }
            public bool IsAmbiguous { get; }
            public List<TextInfo> Candidates { get; }

            public static CandidateResult Empty()
            {
                return new CandidateResult(null, false, new List<TextInfo>());
            }

            public static CandidateResult Single(TextInfo nearest)
            {
                return new CandidateResult(nearest, false, new List<TextInfo> { nearest });
            }

            public static CandidateResult Ambiguous(List<TextInfo> candidates)
            {
                return new CandidateResult(null, true, candidates);
            }
        }

        private sealed class ChangeRecord
        {
            private ChangeRecord(string mark, string status, string oldSpec, string newSpec, string oldLength, string newLength, string note)
            {
                Mark = mark;
                Status = status;
                OldSpec = oldSpec;
                NewSpec = newSpec;
                OldLength = oldLength;
                NewLength = newLength;
                Note = note;
            }

            public string Mark { get; }
            public string Status { get; }
            public string OldSpec { get; }
            public string NewSpec { get; }
            public string OldLength { get; }
            public string NewLength { get; }
            public string Note { get; }

            public static ChangeRecord Updated(string mark, string oldSpec, string newSpec, string oldLength, string newLength)
            {
                return new ChangeRecord(mark, "Updated", oldSpec, newSpec, oldLength, newLength, "");
            }

            public static ChangeRecord Missing(string mark, string note)
            {
                return new ChangeRecord(mark, "Missing", "", "", "", "", note);
            }

            public static ChangeRecord Ambiguous(string mark, CandidateResult specResult, CandidateResult lengthResult)
            {
                var note = new List<string>();
                if (specResult.IsAmbiguous)
                {
                    note.Add("spec");
                }

                if (lengthResult.IsAmbiguous)
                {
                    note.Add("length");
                }

                return new ChangeRecord(mark, "Ambiguous", "", "", "", "", string.Join("+", note));
            }

            public string ToCsv()
            {
                return string.Join(",", Escape(Mark), Escape(Status), Escape(OldSpec), Escape(NewSpec), Escape(OldLength), Escape(NewLength), Escape(Note));
            }

            private static string Escape(string value)
            {
                if (value == null)
                {
                    return "";
                }

                if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
                {
                    var escaped = value.Replace("\"", "\"\"");
                    return $"\"{escaped}\"";
                }

                return value;
            }
        }
    }
}
