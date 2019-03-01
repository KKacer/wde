using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Xml;

namespace jd.Helper.Configuration
{
    /// <summary>
    ///
    /// </summary>
    public class Setting
    {
        public string Name;
        public string Value;

        public Setting(string name, string value)
        {
            Name = name;
            Value=value;
        }
    }

    /// <summary>
    ///
    /// </summary>
    public class Settings:List<Setting>
    {
        static int instanceCount = 0;

        public static int GetInstanceCount()
        {
            return (instanceCount);
        }

        public Setting Add(string name, string value)
        {
            Setting setting = new Setting(name,value);
            Add(setting);
            return setting;
        }

        public Setting GetItemByString(string searchString)
        {
            for (int i = 0; i < Count; i++)
            {
                if (this[i].Name == searchString)
                {
                    return this[i];
                }
            }
            return null;
        }

        public int GetItemIDByString(string searchString)
        {
            for (int i = 0; i < Count; i++)
            {
                if (this[i].Name == searchString)
                {
                    return i;
                }
            }
            return -1;
        }
        //public Setting[] SettingsList;
    }

    /// <summary>
    ///
    /// </summary>
    public class Section
    {
        /// <summary>
        /// Der Name der Sektion
        /// </summary>
        public string Name;
        public string ClassName;
        public Settings settings;

        public Section(string SectionName, string SectionClassName)
        {
            Name = SectionName;
            ClassName = SectionClassName;
            settings = new Settings();
        }
    }

    /// <summary>
    ///
    /// </summary>
    [Serializable()]
    public class Sections:List<Section>
    {
        static int instanceCount = 0;

        //List<Section> SectionList = new List<Section>();

        //Konstruktor
        public Sections() { }

        public static int GetInstanceCount()
        {
            return (instanceCount);
        }

        public Section Add(string sectionName, string SectionClassName)
        {
            instanceCount++;
            Section newSection = new Section(sectionName, SectionClassName);
            Add(newSection);
            //this.SectionList.ad
            //Section section = new Section(sectionName);
            return newSection;
        }

        public Section GetItemByString(string searchString, string searchClassName)
        {
            for (int i = 0; i < Count; i++)
            {
                if ((this[i].Name == searchString) && (this[i].ClassName == searchClassName))
                {
                    return this[i];
                }
            }
            return null;
        }

        public int GetItemIDByString(string searchString, string searchClassName)
        {
            for (int i = 0; i < Count; i++)
            {
                if ((this[i].Name == searchString) && (this[i].ClassName == searchClassName))
                {
                    return i;
                }
            }
            return -1;
        }

    }

    public class ApplicationSettings
    {
        /// <summary>
        /// Speichert den Dateinamen der XML-Datei
        /// </summary>
        private string fileName;

        /// <summary>
        /// Verwaltet die Sektionen
        /// </summary>
        public Sections sections;

        /// <summary>
        /// Konstruktor
        /// </summary>
        /// <param name="fileName">Der Dateiname der XML-Datei</param>
        public ApplicationSettings(string fileName)
        {
            this.fileName = fileName;
            sections = new Sections();

            //Load();
        }

        public string GetNewSectionName()
        {
            int i = 0;

            do
            {
                string tempName = "UserExplorerControl" + i;
                if (sections.GetItemByString(tempName,"OPTIONS") == null)
                {
                    return tempName;
                }

                ++i;

            } while (i < 1000);

            return "";
        }

        public void Save()
        {
            // XmlDocument-Objekt für die Einstellungs-Datei erzeugen
            XmlDocument xmlDoc = new XmlDocument();

            // Skelett der XML-Datei erzeugen
            xmlDoc.LoadXml("<?xml version=\"1.0\" encoding=\"utf-8\" " + "standalone=\"yes\"?><config></config>");
            XmlNode root = xmlDoc.DocumentElement;

            // Alle Sektionen durchgehen und die Einstellungen schreiben
            for (int i = 0; i < sections.Count; i++)
            {
                Section section = sections[i];

                // Element für die Sektion erzeugen und anfügen
                XmlElement sectionElement = xmlDoc.CreateElement("section");

                XmlAttribute attr =xmlDoc.CreateAttribute("name");
                attr.Value = section.Name;
                sectionElement.Attributes.Append(attr);

                XmlAttribute attrSectionClassname = xmlDoc.CreateAttribute("class");
                attrSectionClassname.Value = section.ClassName;
                sectionElement.Attributes.Append(attrSectionClassname);

                //XmlAttribute attrSectionParent = xmlDoc.CreateAttribute("parent");
                //attrSectionParent.Value = section.TabControl;
                //sectionElement.Attributes.Append(attrSectionParent);

                xmlDoc.DocumentElement.AppendChild(sectionElement);

                // Alle Einstellungen der Sektion durchlaufen
                for (int j = 0; j < section.settings.Count; j++)
                {
                    // Einstellungs-Element erzeugen und anfügen
                    XmlElement settingElement = xmlDoc.CreateElement("setting");
                    XmlAttribute attr2 = xmlDoc.CreateAttribute("name");
                    attr2.Value = section.settings[j].Name;
                    settingElement.Attributes.Append(attr2);
                    settingElement.InnerText = section.settings[j].Value;
                    sectionElement.AppendChild(settingElement);
                }
            }

            // Datei speichern
            try
            {
                xmlDoc.Save(fileName);
            }
            catch (IOException ex)
            {
                throw new IOException("Fehler beim Speichern der " +
                   "Konfigurationsdatei '" + fileName + "': " + ex.Message);
            }
            catch (XmlException ex)
            {
                throw new XmlException("Fehler beim Speichern der " +
                   " Konfigurationsdatei '" + fileName + "': " +
                   ex.Message, ex);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        public bool Load()
        {
            if (!File.Exists(fileName))
            {
                return false;
            }

            sections.Clear();


            XmlReader xmlReader = null;

            try
            {
                xmlReader = XmlReader.Create(fileName);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                throw;
            }

            Section currentSection = null;
            string currentSettingName = "";

            while (xmlReader.Read())
            {
                if (xmlReader.NodeType == XmlNodeType.Element)
                {
                    if (xmlReader.Name == "section")
                    {
                        if (xmlReader.AttributeCount > 0)
                        {
                            if (xmlReader.GetAttribute("class") == "FAV")
                            {
                                currentSection = sections.Add(xmlReader.GetAttribute(0), "FAV");
                            }
                            else if (xmlReader.GetAttribute("class") == "OPTIONS")
                            {
                                currentSection = sections.Add(xmlReader.GetAttribute(0), "OPTIONS");
                            }
                            else
                            {
                                currentSection = sections.Add(xmlReader.GetAttribute(0), "");
                            }
                        }
                    }
                    else if (xmlReader.Name == "setting")
                    {
                        if (xmlReader.AttributeCount > 0)
                        {
                            currentSettingName = xmlReader.GetAttribute(0);
                        }
                    }
                }
                else if (xmlReader.NodeType == XmlNodeType.Text)
                {
                    if (currentSection != null)
                    {
                        string settingValue = xmlReader.Value;
                        currentSection.settings.Add(currentSettingName, settingValue);
                        //Console.WriteLine(xmlReader.GetAttribute(0));
                    }
                }
            }
            xmlReader.Close();
            return true;
        }
    }

}
