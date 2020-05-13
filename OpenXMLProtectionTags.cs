public static void LoopThruRecords(string filePath)
        {
            ZipArchive zipArchive = ZipFile.OpenRead(filePath);
            foreach (ZipArchiveEntry archiveEntry in zipArchive.Entries.Where
                (xmlFile => 
                    (xmlFile.FullName.StartsWith(@"xl/worksheets/") || 
                    xmlFile.FullName.Equals(@"xl/workbook.xml")) && Path.GetExtension(filePath + @"\" + xmlFile.FullName) == ".xml"
                ).ToList()
                    )
            {
                string fileExtension = Path.GetExtension(filePath + @"\" + archiveEntry.FullName);
                Debug.Print($"Archive entry path: {archiveEntry.FullName}\n\tFile extension: {fileExtension}\n\tFile size:{archiveEntry.Length.ToString()}");

                XmlTextReader xmlReader = new XmlTextReader(archiveEntry.Open());
                xmlReader.Namespaces = false;
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(xmlReader);

                if (archiveEntry.FullName.Equals(@"xl/workbook.xml"))
                {
                    XmlNode xmlNode = xmlDocument.SelectSingleNode("/workbook/workbookProtection");
                    if (xmlNode != null)
                    {
                        Debug.Print(xmlNode.OuterXml);
                        xmlNode.Attributes.GetNamedItem("lockStructure").InnerXml = "0";
                        Debug.Print(xmlNode.OuterXml);
                    }
                    else
                    {
                        Debug.Print("workbookProtection node does not exist within the current context");
                    }
                }
                else if (archiveEntry.FullName.StartsWith(@"xl/worksheets/"))
                {
                    XmlNode xmlNode = xmlDocument.SelectSingleNode("/worksheet/sheetProtection");
                    if (xmlNode != null)
                    {
                        Debug.Print(xmlNode.OuterXml);
                        xmlNode.RemoveAll();
                        Debug.Print(xmlNode.OuterXml);
                    }
                    else
                    {
                        Debug.Print("sheetProtection node does not exist within the current context");
                    }
                    
                }
                
            }

        }
