using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Xml;

namespace jd.Helper.XXApplication.Configuration
{
    /// <summary>
    /// Manages a setting
    /// </summary>
    public class Setting
    {
        /// <summary>
        /// The name of the setting
        /// </summary>
        public string Name;

        /// <summary>
        /// The value of the setting
        /// </summary>
        public string Value;

        /// <summary>
        /// The default value for setting
        /// </summary>
        public string DefaultValue;

        /// <summary>
        /// Indicates whether the setting
        /// was found in the file
        /// </summary>
        public bool WasInFile;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="name">The name of the setting</param>
        /// <param name="defaultValue">The default value for setting</param>
        public Setting(string name, string defaultValue)
        {
            Name = name;
            DefaultValue = defaultValue;
        }
    }

    /// <summary>
    /// List to save multiple settings
    /// </summary>
    public class Settings : Dictionary<string, Setting>
    {
        /// <summary>
        /// Adds a new setting to the collection
        /// </summary>
        /// <param name="settingName">The name of the setting</param>
        /// <param name="defaultValue">The default value for setting</param>
        public void Add(string settingName, string defaultValue)
        {
            Add(settingName,
                new Setting(settingName, defaultValue));
        }

    }

    /// <summary>
    /// Manages a settings section
    /// </summary>
    public class Section
    {
        /// <summary>
        /// The name of the section
        /// </summary>
        public string Name;

        /// <summary>
        /// The settings of the section
        /// </summary>
        public Settings Settings;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="sectionName">The name of the section</param>
        public Section(string sectionName)
        {
            Name = sectionName;
            Settings = new Settings();
        }
    }

    /// <summary>
    /// List for storing sections
    /// </summary>
    public class Sections : Dictionary<string, Section>
    {
        /// <summary>
        /// Adds a new Section object to the collection
        /// </summary>
        /// <param name="name">The name of the section</param>
        public void Add(string name)
        {
            Add(name, new Section(name));
        }
    }

    /// <summary>
    /// Class for managing configuration data
    /// </summary>
    public class Config
    {
        /// <summary>
        /// Saves the file name of the XML file
        /// </summary>
        private string fileName;

        /// <summary>
        /// Manages the sections
        /// </summary>
        public Sections Sections;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="fileName">Name of the XML file</param>
        public Config(string fileName)
        {
            this.fileName = fileName;
            Sections = new Sections();
        }

        /// <summary>
        /// Leave the configuration data
        /// </summary>
        /// <returns>Returns true if the reading was successful</returns>
        public bool Load()
        {
            // Variable for the return value
            bool returnValue = true;

            // Create an XmlDocument object for the settings file
            XmlDocument xmlDoc = new XmlDocument();

            // load a file
            try
            {
                xmlDoc.Load(fileName);
            }
            catch (IOException ex)
            {
                throw new IOException("Error loading configuration file '" +
                                      fileName + "': " + ex.Message);
            }
            catch (XmlException ex)
            {
                throw new XmlException("Error loading configuration file'" +
                                       fileName + "': " + ex.Message, ex);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }

            // Go through all the sections and read in the settings
            foreach (Section section in Sections.Values)
            {
                // All settings of the section go through
                foreach (Setting setting in section.Settings.Values)
                {
                    // Locate setting in XML document
                    XmlNode settingNode = xmlDoc.SelectSingleNode(
                        "/config/" + section.Name + "/" + setting.Name);
                    if (settingNode != null)
                    {
                        // Setting found
                        setting.Value = settingNode.InnerText;
                        setting.WasInFile = true;
                    }
                    else
                    {
                        // Setting NOT found
                        setting.Value = setting.DefaultValue;
                        setting.WasInFile = false;
                        returnValue = false;
                    }
                }
            }

            // Report result
            return returnValue;
        }

        /// <summary>
        /// Saves the configuration data
        /// </summary>
        public void Save()
        {
            // Create an XmlDocument object for the settings file
            XmlDocument xmlDoc = new XmlDocument();

            // Create skeleton of the XML file
            xmlDoc.LoadXml("<?xml version=\"1.0\" encoding=\"utf-8\" " +
                           "standalone=\"yes\"?><config></config>");

            // Go through all the sections and write the settings
            foreach (Section section in Sections.Values)
            {
                // Create and attach an element for the section
                XmlElement sectionElement = xmlDoc.CreateElement(section.Name);
                xmlDoc.DocumentElement.AppendChild(sectionElement);

                // All settings of the section go through
                foreach (Setting setting in section.Settings.Values)
                {
                    // Create and attach a settings item
                    XmlElement settingElement =
                        xmlDoc.CreateElement(setting.Name);
                    settingElement.InnerText = setting.Value;
                    sectionElement.AppendChild(settingElement);
                }
            }

            // save file
            try
            {
                xmlDoc.Save(fileName);
            }
            catch (IOException ex)
            {
                throw new IOException("Error saving the " +
                                      "configuration file '" + fileName + "': " + ex.Message);
            }
            catch (XmlException ex)
            {
                throw new XmlException("Error saving the " +
                                       "configuration file '"+ fileName + "': " +
                                       ex.Message, ex);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }
    }
}
