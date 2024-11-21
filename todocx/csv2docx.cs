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
                    if (text.Text.Contains("{{StartTime}}"))
                    {
                        text.Text = text.Text.Replace("{{StartTime}}", startTime);
                    }
                    if (text.Text.Contains("{{EndTime}}"))
                    {
                        text.Text = text.Text.Replace("{{EndTime}}", endTime);
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
                    if (text.Text.Contains("{{ExperimentName}}"))
                    {
                        text.Text = text.Text.Replace("{{ExperimentName}}", experimentName);
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
                        if (text.Text.Contains("{{" + key + "}}"))
                        {
                            text.Text = text.Text.Replace("{{" + key + "}}", seriesData[key]);
                        }
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

                // Find the table in the document
                var table = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().FirstOrDefault();

                // Get the first row of the table (header row)
                var headerRow = table.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableRow>().First();

                // Create a new row for each data row in the CSV file
                foreach (var rowData in csvData)
                {
                    // Clone the header row to create a new row
                    var newRow = (DocumentFormat.OpenXml.Wordprocessing.TableRow)headerRow.CloneNode(true);

                    // Find and replace the placeholders in the new row with the data values
                    foreach (var cell in newRow.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>())
                    {
                        foreach (var text in cell.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                        {
                            foreach (var key in rowData.Keys)
                            {
                                if (text.Text.Contains("{{" + key + "}}"))
                                {
                                    text.Text = text.Text.Replace("{{" + key + "}}", rowData[key]);
                                }
                            }
                        }
                    }

                    // Add the new row to the table
                    table.AppendChild(newRow);
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
            /// <Experiment>
            ///     <StartTime>2022-01-01 12:00:00</StartTime>
            ///     <EndTime>2022-01-01 13:00:00</EndTime>
            ///     <ExperimentName>Sample Experiment</ExperimentName>
            ///     <ElasticModulus>100 GPa</ElasticModulus>
            ///     <Density>2.7 g/cm^3</Density>
            ///     <MaxStress>200 MPa</MaxStress>
            ///     <RatioOfStress>0.5</RatioOfStress>
            ///     <CycleCount>1000</CycleCount>
            ///     <BottomAmplitude>10 mm</BottomAmplitude>
            /// </Experiment>
            var xmlDoc = new System.Xml.XmlDocument();
            xmlDoc.LoadXml(xmlData);

            // Extract the data from the XML object
            string startTime = xmlDoc.SelectSingleNode("/Experiment/StartTime").InnerText;
            string endTime = xmlDoc.SelectSingleNode("/Experiment/EndTime").InnerText;
            string experimentName = xmlDoc.SelectSingleNode("/Experiment/ExperimentName").InnerText;
            double elasticModulus = double.Parse(xmlDoc.SelectSingleNode("/Experiment/ElasticModulus").InnerText.Replace("GPa", ""));
            double density = double.Parse(xmlDoc.SelectSingleNode("/Experiment/Density").InnerText.Replace("g/cm^3", ""));
            double maxStress = double.Parse(xmlDoc.SelectSingleNode("/Experiment/MaxStress").InnerText.Replace("MPa", ""));
            double ratioOfStress = double.Parse(xmlDoc.SelectSingleNode("/Experiment/RatioOfStress").InnerText);
            long cycleCount = long.Parse(xmlDoc.SelectSingleNode("/Experiment/CycleCount").InnerText);
            double bottomAmplitude = double.Parse(xmlDoc.SelectSingleNode("/Experiment/BottomAmplitude").InnerText.Replace("mm", ""));
        
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
                { "BottomAmplitude", bottomAmplitude.ToString() }
            };
        }

        /// <summary>
        /// The Main method is the entry point for the CSV to DOCX conversion process.
        /// </summary>
        /// <param name="xmlPath">The path to the xmlPath contain stattime endtime experiment name and seria data.</param>
        /// <param name="csvPath">The path to the CSV file.</param> 
        /// <param name="templatePath">The path to the template DOCX file.</param>
        /// <param name="outputPath">The path to the output DOCX file.</param>
        public void GenerateDocx(string xmlPath, string csvPath, string templatePath, string outputPath)
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

            // Write the series data to the template DOCX file
            WriteSeriesData(templatePath, xmlData);

            // Write the CSV data to the output DOCX file
            WriteCsvDataToDocx(csvPath, outputPath);
        }
    }
}
