using LevDan.Exif;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using ExcelApp = Microsoft.Office.Interop.Excel;

namespace GPSExif
{
    class Program
    {
        const string NO_ARGS = "Please enter arguments for source path and results path. Use GPSExif.exe -h for help";
        const string HELP = @"Usage is: GPSExif.exe -s c:\pictures folder\ -d c:\report output path -k c:\kml output path  (folders in double quotes if there are spaces in the names)";
        const string PROCESSING = "Processing photos from {0} to {1}";

        public static ExcelApp.Application app = null;
        public static ExcelApp.Workbook workbook = null;
        public static ExcelApp.Worksheet gpsWorksheet = null;
        public static ExcelApp.Worksheet audioWorksheet = null;
        public static ExcelApp.Worksheet videoWorksheet = null;
        public static ExcelApp.Worksheet documentWorksheet = null;

        public static void CreateHeaders(int row, int col, string htext, ExcelApp.Worksheet worksheet)
        {
            worksheet.Cells[row, col] = htext;
            worksheet.Cells[row, col] = "Bold";
        }
        public static void addData(int row, int col, string data, ExcelApp.Worksheet worksheet)
        {
            worksheet.Cells[row, col] = data;
        }

        static void Main(string[] args)
        {
            app = new ExcelApp.Application();
            app.DisplayAlerts = false;
            app.Visible = false;
            int row = 2;

            workbook = app.Workbooks.Add(1);
            gpsWorksheet = workbook.Sheets.Add();
            //audioWorksheet = workbook.Sheets.Add();
            //videoWorksheet = workbook.Sheets.Add();
            //documentWorksheet = workbook.Sheets.Add();

            gpsWorksheet.Name = "GPS";
            //audioWorksheet.Name = "Audio";
            //videoWorksheet.Name = "Video";
            //documentWorksheet.Name = "Document";

            //creates the main header
            CreateHeaders(1, 1, "Filename", gpsWorksheet);
            CreateHeaders(1, 2, "Latitude", gpsWorksheet);
            CreateHeaders(1, 3, "Longitude", gpsWorksheet);
            CreateHeaders(1, 4, "Altitude", gpsWorksheet);
            CreateHeaders(1, 5, "Make", gpsWorksheet);
            CreateHeaders(1, 6, "Model", gpsWorksheet);
            CreateHeaders(1, 7, "Modify Date Time", gpsWorksheet);
            CreateHeaders(1, 8, "GPS Time (Atomic Clock)", gpsWorksheet);
            CreateHeaders(1, 9, "Date Time Original", gpsWorksheet);
            CreateHeaders(1, 10, "Date Time Digitized", gpsWorksheet);
            CreateHeaders(1, 11, "GPS Date Stamp", gpsWorksheet);

            //StringBuilder sb = new StringBuilder();
            double latitude = 0;
            double longitude = 0;
            string altitude = string.Empty;
            string modifyDateTime = string.Empty;
            string gpsTimeStamp = string.Empty;
            string dateTimeOriginal = string.Empty;
            string dateTimeDigitized = string.Empty;
            string gpsDateStamp = string.Empty;
            string make = string.Empty;
            string model = string.Empty;

            List<string> photos = new List<string>();

            for (int i = 0; i < args.Length; i++)
            {
                Console.WriteLine(string.Format("ARGS{0}={1}", i, args[i].ToString()));
            }

            if (args.Length == 0)
            {
                Console.WriteLine(NO_ARGS);
            }
            else
            {
                if (args[0].ToString() == "-h")
                {
                    Console.WriteLine(HELP);
                }

                Console.WriteLine("ARGS length = " + args.Length);


                if (args[0].ToString() == "-s" && args[2].ToString() == "-d" && args[4].ToString() == "-k")
                {

                    Console.WriteLine(string.Format(PROCESSING, args[1].ToString(), args[3].ToString()));
                    DirectoryInfo source = new DirectoryInfo(args[1].ToString());
                    DirectoryInfo excelReportDestination = new DirectoryInfo(args[3].ToString());
                    DirectoryInfo kmlReportDestination = new DirectoryInfo(args[5].ToString());

                    foreach (FileInfo path in source.GetFiles("*", SearchOption.AllDirectories))
                    {
                        try
                        {
                            ExifGPSLatLonTagCollection exif = new ExifGPSLatLonTagCollection(path.FullName);

                            if (exif.Count() >= 3)//datetime, lat, long
                            {
                                foreach (ExifTag tag in exif)
                                {
                                    //Console.WriteLine(tag.FieldName);
                                    string latRef = string.Empty;
                                    string lonRef = string.Empty;

                                    foreach (ExifTag tag2 in exif)
                                    {
                                        switch (tag2.FieldName)
                                        {
                                            case "GPSLatitudeRef":
                                                {
                                                    latRef = tag2.Value;
                                                    break;
                                                }
                                            case "GPSLongitudeRef":
                                                {
                                                    lonRef = tag2.Value;
                                                    break;
                                                }
                                        }
                                    }
                                    switch (tag.FieldName)
                                    {
                                        case "GPSLatitude":
                                            {
                                                if (!string.IsNullOrEmpty(latRef))
                                                {
                                                    latitude = Utilities.GPS.GetLatLonFromDMS(latRef.Substring(0, 1) + tag.Value);
                                                }
                                                latitude = Utilities.GPS.GetLatLonFromDMS(tag.Value);
                                                break;
                                            }
                                        case "GPSLongitude":
                                            {
                                                if (!string.IsNullOrEmpty(lonRef))
                                                {
                                                    longitude = Utilities.GPS.GetLatLonFromDMS(lonRef.Substring(0, 1) + tag.Value);
                                                }
                                                longitude = Utilities.GPS.GetLatLonFromDMS(tag.Value);
                                                break;
                                            }
                                        case "GPSAltitude":
                                            {
                                                altitude = tag.Value;
                                                break;
                                            }
                                        case "DateTime":
                                            {
                                                modifyDateTime = tag.Value;
                                                break;
                                            }
                                        case "Make":
                                            {
                                                make = tag.Value;
                                                break;
                                            }
                                        case "Model":
                                            {
                                                model = tag.Value;
                                                break;
                                            }
                                        case "DateTimeOriginal":
                                            {
                                                dateTimeOriginal = tag.Value;
                                                break;
                                            }
                                        case "DateTimeDigitized":
                                            {
                                                dateTimeDigitized = tag.Value;
                                                break;
                                            }
                                        case "GPSDateStamp":
                                            {
                                                gpsDateStamp = tag.Value;
                                                break;
                                            }
                                        case "GPSTimeStamp":
                                            {
                                                gpsTimeStamp = tag.Value;
                                                break;
                                            }
                                    }
                                }

                                if (latitude > 0 && longitude > 0)
                                {
                                    addData(row, 1, Path.GetFileName(path.FullName).ToString(), gpsWorksheet);
                                    addData(row, 2, latitude.ToString(), gpsWorksheet);
                                    addData(row, 3, longitude.ToString(), gpsWorksheet);
                                    addData(row, 4, altitude.ToString(), gpsWorksheet);
                                    addData(row, 5, make.ToString(), gpsWorksheet);
                                    addData(row, 6, model.ToString(), gpsWorksheet);
                                    addData(row, 7, modifyDateTime.ToString(), gpsWorksheet);
                                    addData(row, 8, dateTimeOriginal.ToString(), gpsWorksheet);
                                    addData(row, 9, dateTimeDigitized.ToString(), gpsWorksheet);
                                    addData(row, 10, gpsDateStamp.ToString(), gpsWorksheet);
                                    addData(row, 11, gpsTimeStamp.ToString(), gpsWorksheet);
                                    row++;
                                }

                                photos.Add(longitude + "," + latitude + "," + altitude + "," + modifyDateTime + "," + Path.GetFileName(path.FullName) + "," + make + "," + model);
                                latitude = 0;
                                longitude = 0;
                                altitude = string.Empty;
                                modifyDateTime = string.Empty;
                                make = string.Empty;
                                model = string.Empty;
                            }
                        }
                        catch (Exception) { }

                        if (photos.Count > 0)
                        {
                            string path3 = Path.Combine(kmlReportDestination.FullName, "gps_exif_report.kml");
                            KML.Create(photos, path3);

                            //string path2 = Path.Combine(excelReportDestination.FullName, "report.txt");
                            //File.WriteAllText(path2, sb.ToString());
                        }

                    }
                    gpsWorksheet.Columns.AutoFit();
                    workbook.SaveAs(Path.Combine(excelReportDestination.FullName, "GPS Report.xlsx"));
                    workbook.Close();
                }
                else
                {
                    Console.WriteLine(HELP);
                }
            }

            //Console.WriteLine("Done processing files");
        }
        /// <summary>
        /// Creates a Point and Placemark and prints the resultant KML on to the console.
        /// </summary>
        public static class KML
        {
            public static void Create(List<string> points, string path)
            {
                //photos.Add(longitude + "," + latitude + "," + altitude + "," + dateTime + "," + Path.GetFileName(path.FullName) + "," + make + "," + model);
                StringBuilder sb = new StringBuilder();

                using (XmlWriter writer = XmlWriter.Create(path))
                {
                    int i = 1;
                    writer.WriteStartElement("Document");
                    writer.WriteElementString("name", "photos.xml");
                    writer.WriteElementString("open", "1");

                    writer.WriteStartElement("Style");
                    writer.WriteStartElement("LabelStyle");
                    writer.WriteElementString("color", "ff0000cc");
                    writer.WriteEndElement();//LabelStyle
                    writer.WriteEndElement();//Style

                    foreach (string p in points)
                    {
                        string[] splitUp = p.Split(',');
                        sb.Append("Latitude: " + splitUp[0].ToString());
                        sb.Append(Environment.NewLine);
                        sb.Append("Longitude: " + splitUp[1].ToString());
                        sb.Append(Environment.NewLine);
                        sb.Append("Altitude: " + splitUp[2].ToString());
                        sb.Append(Environment.NewLine);
                        sb.Append("Date Time: " + splitUp[3].ToString());
                        sb.Append(Environment.NewLine);
                        sb.Append("File Name: " + splitUp[4].ToString());
                        sb.Append(Environment.NewLine);
                        sb.Append("Make: " + splitUp[5].ToString());
                        sb.Append(Environment.NewLine);
                        sb.Append("Model: " + splitUp[6].ToString());

                        if (splitUp[0].ToString() != "0" && splitUp[1].ToString() != "0")
                        {
                            writer.WriteStartElement("Placemark");
                            writer.WriteElementString("description", sb.ToString());
                            writer.WriteElementString("name", splitUp[4].ToString());
                            writer.WriteStartElement("Point");
                            writer.WriteElementString("coordinates", splitUp[0].ToString() + "," + splitUp[1].ToString());
                            writer.WriteEndElement();//Point
                            writer.WriteEndElement();//Placemark
                            i++;
                        }
                        sb.Clear();
                    }

                    writer.WriteEndElement();//Document
                    writer.Flush();
                }
            }
        }

        public static void CreateHeader(int row, int col, string htext, Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            worksheet.Cells[row, col] = htext;
        }
        public static void AddData(int row, int col, string data, Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            worksheet.Cells[row, col] = data;
        }
    }
}
