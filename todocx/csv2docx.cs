/// <summary>
/// The csv2docx.cs file contains the implementation for converting CSV (Comma-Separated Values) files into DOCX (Microsoft Word Document) format.
/// This file includes methods for reading CSV data, processing it, and generating a DOCX document with the corresponding content.
/// </summary> 
/// <remarks>
/// This file is part of the todocx project, which aims to automate the creation of Word documents based on CSV data and a template file.
/// </remarks>

using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace todocx
{
    /// <summary>
    /// The Csv2Docx class provides methods for converting CSV files to DOCX format.
    /// 1. write start time and end time to template docx file
    /// 2. write experiment name to template docx file
    /// 3. write a series data, eg: Elastic modulus，density，max stress，ratio of stress, cycle count, Bottom amplitude to template docx file
    /// 4. read csv data and write to docx 
    /// </summary>
    public class Csv2Docx
    {
        /// <summary>
        /// The WriteTimeInfo method writes the start time and end time to the specified DOCX file.
        /// </summary>
        /// <param name="docxPath">The path to the DOCX file.</param>
        /// <param name="startTime">The start time to be written to the template DOCX file.</param>
        /// <param name="endTime">The end time to be written to the template DOCX file.</param>
        private void WriteTimeInfo(string docxPath, string startTime, string endTime)
        {
            // Open the template DOCX file
            using (WordprocessingDocument doc = WordprocessingDocument.Open(docxPath, true))
            {
                // Access the main document part
                var mainPart = doc.MainDocumentPart;

                // Get the document body
                var body = mainPart.Document.Body;

                // Find and replace the placeholders for start time and end time
                foreach (var text in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                {
                    if (text.Text.Contains("StartTime"))
                    {
                        text.Text = text.Text.Replace("StartTime", startTime);
                    }
                    if (text.Text.Contains("EndTime"))
                    {
                        text.Text = text.Text.Replace("EndTime", endTime);
                    }
                }

                // Save the changes
                mainPart.Document.Save();
            }
        }

        /// <summary>
        /// The WriteExperimentName method writes the experiment name to the specified DOCX file.
        /// </summary>
        /// <param name="docxPath">The path to the DOCX file.</param>
        /// <param name="experimentName">The experiment name to be written to the template DOCX file.</param>
        private void WriteExperimentName(string docxPath, string experimentName)
        {
            // Open the template DOCX file
            using (WordprocessingDocument doc = WordprocessingDocument.Open(docxPath, true))
            {
                // Access the main document part
                var mainPart = doc.MainDocumentPart;

                // Get the document body
                var body = mainPart.Document.Body;

                // Find and replace the placeholder for experiment name
                foreach (var text in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                {
                    if (text.Text.Contains("ExperimentName"))
                    {
                        text.Text = text.Text.Replace("ExperimentName", experimentName);
                    }
                }

                // Save the changes
                mainPart.Document.Save();
            }
        }

        /// <summary>
        /// The WriteSeriesData method writes a series of data to the specified DOCX file.
        /// </summary>
        /// <param name="docxPath">The path to the DOCX file.</param>
        /// <param name="seriesData">The series data to be written to the template DOCX file.</param>
        private void WriteSeriesData(string docxPath, Dictionary<string, string> seriesData)
        {
            // Open the template DOCX file
            using (WordprocessingDocument doc = WordprocessingDocument.Open(docxPath, true))
            {
                // Access the main document part
                var mainPart = doc.MainDocumentPart;

                // Get the document body
                var body = mainPart.Document.Body;

                // Find and replace the placeholders for series data
                foreach (var text in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                {
                    foreach (var key in seriesData.Keys)
                    {
                        if (text.Text.Contains(key))
                        {
                            text.Text = text.Text.Replace(key, seriesData[key]);
                        }
                    }
                }

                // Save the changes
                mainPart.Document.Save();
            }
        }

        /// <summary>
        /// Write IntermittentExp method writes the intermittent experiment data to the specified DOCX file.
        /// </summary>
        /// <param name="docxPath">The path to the DOCX file.</param>
        /// <param name="intermittentExp">The intermittent experiment data to be written to the template DOCX file.</param>
        /// <param name="excitationtime">The excitation time to be written to the template DOCX file.</param>
        /// <param name="intervaltime">The interval time to be written to the template DOCX file.</param>
        private void WriteIntermittentExp(string docxPath, string intermittentExp, string excitationtime, string intervaltime)
        {
            // Open the template DOCX file
            using (WordprocessingDocument doc = WordprocessingDocument.Open(docxPath, true))
            {
                // Access the main document part
                var mainPart = doc.MainDocumentPart;

                // Get the document body
                var body = mainPart.Document.Body;

                // Find and replace the placeholders for intermittent experiment data
                foreach (var text in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                {
                    if (text.Text.Contains("IntermittentExp"))
                    {
                        text.Text = text.Text.Replace("IntermittentExp", intermittentExp);
                    }
                    if (text.Text.Contains("ExcitationTime"))
                    {
                        text.Text = text.Text.Replace("ExcitationTime", (excitationtime=="0")?"—":excitationtime);
                    }
                    if (text.Text.Contains("IntervalTime"))
                    {
                        text.Text = text.Text.Replace("IntervalTime", (intervaltime=="0")?"—":intervaltime);
                    }
                }

                // Save the changes
                mainPart.Document.Save();
            }
        }
        /// <summary>
        /// The ReadCsvData method reads data from a CSV file and returns it as a list of dictionaries.
        /// Each dictionary represents a row in the CSV file, with column headers as keys and values as values.
        /// </summary>
        /// <param name="csvPath">The path to the CSV file.</param>
        /// <returns>A list of dictionaries representing the data from the CSV file.</returns>
        private List<Dictionary<string, string>> ReadCsvData(string csvPath)
        {
            List<Dictionary<string, string>> csvData = new List<Dictionary<string, string>>();

            // Read the CSV file line by line
            using (var reader = new StreamReader(csvPath))
            {
                // Read the header row to get the column names
                var headerLine = reader.ReadLine();
                var columnNames = headerLine.Split(',');

                // Read the data rows
                while (!reader.EndOfStream)
                {
                    var dataLine = reader.ReadLine();
                    var dataValues = dataLine.Split(',');

                    // Create a dictionary for the row data
                    var rowData = new Dictionary<string, string>();
                    for (int i = 0; i < columnNames.Length; i++)
                    {
                        rowData.Add(columnNames[i], dataValues[i]);
                    }

                    // Add the row data to the list
                    csvData.Add(rowData);
                }
            }

            return csvData;
        }

        /// <summary>
        /// The WriteCsvDataToDocx method reads data from a CSV file and writes it to the specified DOCX file.
        /// replace tag ExperienceDataList with csv data in the docx file
        /// </summary>
        /// <param name="csvPath">The path to the CSV file.</param>
        /// <param name="docxPath">The path to the DOCX file.</param>
        private void WriteCsvDataToDocx(string csvPath, string docxPath)
        {
            // Read the CSV data
            var csvData = ReadCsvData(csvPath);

            // Open the template DOCX file
            using (WordprocessingDocument doc = WordprocessingDocument.Open(docxPath, true))
            {
                // Access the main document part
                var mainPart = doc.MainDocumentPart;

                // Get the document body
                var body = mainPart.Document.Body;

                // Find the tag for the ExperienceDataList
                var tag = "ExperienceDataList";
                foreach (var text in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                {
                    if (text.Text.Contains(tag))
                    {
                        // Make the tag ExperienceDataList paragraph center alignment
                        var paragraph = text.Parent;
                        var paragraphProperties = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties(
                            new DocumentFormat.OpenXml.Wordprocessing.Justification { Val = DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Center }
                        );
                        // Remove the tag from the text
                        text.Text = text.Text.Replace(tag, "");
                        // create a table to store the csv data after the tag
                        var table = new DocumentFormat.OpenXml.Wordprocessing.Table();
                        // create the table properties
                        // make the table border visible
                        var tableProperties = new DocumentFormat.OpenXml.Wordprocessing.TableProperties(
                            new DocumentFormat.OpenXml.Wordprocessing.TableBorders(
                                new DocumentFormat.OpenXml.Wordprocessing.TopBorder
                                {
                                    Val = new EnumValue<DocumentFormat.OpenXml.Wordprocessing.BorderValues>(DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single),
                                    Size = 12
                                },
                                new DocumentFormat.OpenXml.Wordprocessing.BottomBorder
                                {
                                    Val = new EnumValue<DocumentFormat.OpenXml.Wordprocessing.BorderValues>(DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single),
                                    Size = 12
                                },
                                new DocumentFormat.OpenXml.Wordprocessing.LeftBorder
                                {
                                    Val = new EnumValue<DocumentFormat.OpenXml.Wordprocessing.BorderValues>(DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single),
                                    Size = 12
                                },
                                new DocumentFormat.OpenXml.Wordprocessing.RightBorder
                                {
                                    Val = new EnumValue<DocumentFormat.OpenXml.Wordprocessing.BorderValues>(DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single),
                                    Size = 12
                                },
                                new DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder
                                {
                                    Val = new EnumValue<DocumentFormat.OpenXml.Wordprocessing.BorderValues>(DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single),
                                    Size = 12
                                },
                                new DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder
                                {
                                    Val = new EnumValue<DocumentFormat.OpenXml.Wordprocessing.BorderValues>(DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single),
                                    Size = 12
                                }
                            )
                        );
                        table.AppendChild(tableProperties);
                        // make the table middle align in the page
                        var tableJustification = new DocumentFormat.OpenXml.Wordprocessing.TableJustification { Val = DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues.Center };
                        table.AppendChild(tableJustification);
                        // make the table width 100% of the page
                        var tableWidth = new DocumentFormat.OpenXml.Wordprocessing.TableWidth { Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Pct, Width = "100%" };
                        table.AppendChild(tableWidth);
                        // create a table row for the header
                        var headerRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
                        foreach (var columnName in csvData.First().Keys)
                        {
                            var headerCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                            headerCell.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(columnName))));
                            headerRow.Append(headerCell);
                        }
                        table.Append(headerRow);
                        // create a table row for each data row
                        foreach (var dataRow in csvData)
                        {
                            var dataRowElement = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
                            foreach (var dataValue in dataRow.Values)
                            {
                                var dataCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                                dataCell.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(dataValue))));
                                dataRowElement.Append(dataCell);
                            }
                            table.Append(dataRowElement);
                        }
                        // insert the table after the tag
                        text.Parent.InsertAfterSelf(table);
                    }
                }
                // Save the changes
                mainPart.Document.Save();
            }
        }

        private Dictionary<string, string> ReadXmlData(string xmlPath)
        {
            // Read the JSON data from the file
            string xmlData = File.ReadAllText(xmlPath);

            // Parse the xmlData to extract the relevant information
            // xml data format e.g
            /// <ExperimentReport>
            ///     <StartTime>2022-01-01 12:00:00</StartTime>
            ///     <EndTime>2022-01-01 13:00:00</EndTime>
            ///     <ExperimentName>Sample ExperimentReport</ExperimentName>
            ///     <ElasticModulus>100 GPa</ElasticModulus>
            ///     <Density>2.7 g/cm^3</Density>
            ///     <MaxStress>200 MPa</MaxStress>
            ///     <RatioOfStress>0.5</RatioOfStress>
            ///     <CycleCount>1000</CycleCount>
            ///     <BottomAmplitude>10 mm</BottomAmplitude>
            ///     <IntermittentExp>1</IntermittentExp>
            ///     <ExcitationTime>100ms</ExcitationTime>
            ///     <IntervalTime>100ms</IntervalTime>
            ///     <ExpMode>0</ExpMode>
            /// </ExperimentReport>
            var xmlDoc = new System.Xml.XmlDocument();
            xmlDoc.LoadXml(xmlData);

            // Extract the data from the XML object
            string startTime = xmlDoc.SelectSingleNode("/ExperimentReport/StartTime").InnerText;
            string endTime = xmlDoc.SelectSingleNode("/ExperimentReport/EndTime").InnerText;
            string experimentName = xmlDoc.SelectSingleNode("/ExperimentReport/ExperimentName").InnerText;
            double elasticModulus = double.Parse(xmlDoc.SelectSingleNode("/ExperimentReport/ElasticModulus").InnerText.Replace("GPa", ""));
            double density = double.Parse(xmlDoc.SelectSingleNode("/ExperimentReport/Density").InnerText.Replace("g/cm^3", ""));
            double maxStress = double.Parse(xmlDoc.SelectSingleNode("/ExperimentReport/MaxStress").InnerText.Replace("MPa", ""));
            double ratioOfStress = double.Parse(xmlDoc.SelectSingleNode("/ExperimentReport/RatioOfStress").InnerText);
            long cycleCount = long.Parse(xmlDoc.SelectSingleNode("/ExperimentReport/CycleCount").InnerText);
            double bottomAmplitude = double.Parse(xmlDoc.SelectSingleNode("/ExperimentReport/BottomAmplitude").InnerText.Replace("mm", ""));
            int intermittentExp = int.Parse(xmlDoc.SelectSingleNode("/ExperimentReport/IntermittentExp").InnerText);
            string excitationtime = xmlDoc.SelectSingleNode("/ExperimentReport/ExcitationTime").InnerText;
            string intervaltime = xmlDoc.SelectSingleNode("/ExperimentReport/IntervalTime").InnerText;
            int expmode = int.Parse(xmlDoc.SelectSingleNode("/ExperimentReport/ExpMode").InnerText);
        
            return new Dictionary<string, string>
            {
                { "StartTime", startTime },
                { "EndTime", endTime },
                { "ExperimentName", experimentName },
                { "ElasticModulus", elasticModulus.ToString() },
                { "Density", density.ToString() },
                { "MaxStress", maxStress.ToString() },
                { "RatioOfStress", ratioOfStress.ToString() },
                { "CycleCount", cycleCount.ToString() },
                { "BottomAmplitude", bottomAmplitude.ToString() },
                { "IntermittentExp", intermittentExp.ToString() },
                { "ExcitationTime", excitationtime },
                { "IntervalTime", intervaltime },
                { "ExpMode", expmode.ToString() },
            };
        }

        /// <summary>
        /// The Main method is the entry point for the CSV to DOCX conversion process.
        /// </summary>
        /// <param name="xmlPath">The path to the xmlPath contain stattime endtime experiment name and seria data.</param>
        /// <param name="csvPath">The path to the CSV file.</param> 
        /// <param name="templatePath">The path to the template DOCX file.</param>
        public int GenerateDocx(string xmlPath, string csvPath, string templatePath)
        {
            // Read xmlPath data
            var xmlData = ReadXmlData(xmlPath);

            // Extract the data from the JSON object
            string startTime = xmlData["StartTime"];
            string endTime = xmlData["EndTime"];
            string experimentName = xmlData["ExperimentName"];

            // Write the time info to the template DOCX file
            WriteTimeInfo(templatePath, startTime, endTime);

            // Write the experiment name to the template DOCX file
            WriteExperimentName(templatePath, experimentName);

            string intermittentExp = xmlData["IntermittentExp"];
            string excitationtime = xmlData["ExcitationTime"];
            string intervaltime = xmlData["IntervalTime"];
            // write intermittent experiment data to the template DOCX file
            WriteIntermittentExp(templatePath, intermittentExp, excitationtime, intervaltime);

            // Write the series data to the template DOCX file
            WriteSeriesData(templatePath, xmlData);

            // Write the CSV data to the output DOCX file
            WriteCsvDataToDocx(csvPath, templatePath);

            return 0;
        }
    }
}
