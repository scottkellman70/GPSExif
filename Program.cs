using LevDan.Exif;
using SharpKml.Base;
using SharpKml.Dom;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Directory = System.IO.Directory;

namespace GPSExif
{
    class Program
    {
        const string NO_ARGS = "Please enter arguments for source path and results path. Use GPSExif.exe -h for help";
        const string HELP = @"Usage is: GPSExif.exe -s c:\pictures folder\ -d c:\output folder\ (folders in double quotes if there are spaces in the names)";
        const string PROCESSING = "Processing photos from {0} to {1}";
        const string NO_DIRECTORY = "Directory {0} does not exist.";

        static void Main(string[] args)
        {
            StringBuilder sb = new StringBuilder();
            double latitude = 0;
            double longitude = 0;
            string altitude = string.Empty;
            string dateTime = string.Empty;
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


                if (args[0].ToString() == "-s" && args[2].ToString() == "-d")
                {
                    if (!Directory.Exists(args[1].ToString()) || !Directory.Exists(args[3].ToString()))
                    {
                        if (!Directory.Exists(args[1].ToString()))
                        {
                            Console.WriteLine(string.Format(NO_DIRECTORY, args[1].ToString()));
                        }
                        else
                        {
                            Console.WriteLine(string.Format(NO_DIRECTORY, args[3].ToString()));
                        }
                    }
                    else
                    {
                        Console.WriteLine(string.Format(PROCESSING, args[1].ToString(), args[3].ToString()));
                        DirectoryInfo source = new DirectoryInfo(args[1].ToString());
                        DirectoryInfo destination = new DirectoryInfo(args[3].ToString());

                        foreach (FileInfo path in source.GetFiles("*", SearchOption.AllDirectories))
                        {
                            try
                            {
                                ExifGPSLatLonTagCollection exif = new ExifGPSLatLonTagCollection(path.FullName);

                                if (exif.Count() >= 3)//datetime, lat, long
                                {
                                    /*
                                     public GPSLatLonTags()
                                        {
                                            this.Add(0x2, new ExifTag(0x2, "GPSLatitude", "Latitude"));
                                            this.Add(0x4, new ExifTag(0x4, "GPSLongitude", "Longitude"));
                                            this.Add(0x132, new ExifTag(0x132, "DateTime", "File change date and time"));
                                            this.Add(0x6, new ExifTag(0x6, "GPSAltitude", "Altitude"));
                                            this.Add(0x10E, new ExifTag(0x10E, "ImageDescription", "Image title"));
                                            this.Add(0x10F, new ExifTag(0x10F, "Make", "Image input equipment manufacturer"));
                                            this.Add(0x110, new ExifTag(0x110, "Model", "Image input equipment model"));
                                        }
                                     */
                                    foreach (ExifTag tag in exif)
                                    {
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
                                                    dateTime = tag.Value;
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
                                        }
                                    }
                                    sb.Append(Path.GetFileName(path.FullName) + "," + latitude + "," + longitude + "," + altitude + "," + make + "," + model);
                                    sb.Append(Environment.NewLine);
                                    photos.Add(longitude + "," + latitude + "," + altitude + "," + dateTime + "," + Path.GetFileName(path.FullName) + "," + make + "," + model);
                                    latitude = 0;
                                    longitude = 0;
                                    altitude = string.Empty;
                                    dateTime = string.Empty;
                                    make = string.Empty;
                                    model = string.Empty;
                                }
                            }
                            catch (Exception) { }
                        }

                        if (photos.Count > 0)
                        {
                            string path3 = Path.Combine(destination.FullName, "photos.kml");
                            KML.Create(photos, path3);

                            string path2 = Path.Combine(destination.FullName, "report.txt");
                            File.WriteAllText(path2, sb.ToString());
                        }

                        Console.WriteLine("Done processing files");
                    }
                }
                else
                {
                    Console.WriteLine(HELP);
                }
            }
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
        public static class CreateIconStyle
        {
            public static void Run()
            {
                // Create the style first
                var style = new Style();
                style.Id = "randomColorIcon";
                style.Icon = new IconStyle();
                style.Icon.Color = new Color32(255, 0, 255, 0);
                style.Icon.ColorMode = ColorMode.Random;
                style.Icon.Icon = new IconStyle.IconLink(new Uri("http://maps.google.com/mapfiles/kml/pal3/icon21.png"));
                style.Icon.Scale = 1.1;

                // Now create the object to apply the style to
                var placemark = new Placemark();
                placemark.Name = "IconStyle.kml";
                placemark.StyleUrl = new Uri("#randomColorIcon", UriKind.Relative);
                placemark.Geometry = new Point
                {
                    Coordinate = new Vector(37.831145, -122.36868)
                };

                // Package it all together...
                var document = new Document();
                document.AddFeature(placemark);
                document.AddStyle(style);

                // And display the result
                var serializer = new Serializer();
                serializer.Serialize(document);
                Console.WriteLine(serializer.Xml);
            }
        }
    }
}
