using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;

namespace SolEolImportExport.Domain
{
    [XmlRoot("eExact")]
    public class Topics
    {
        [XmlArray("Topics"), XmlArrayItem("Topic", typeof(Topic))]
        public List<Topic> TopicsList { get; set; }

        [XmlIgnore]
        public string[] TopicCodes
        {
            get
            {
                var topicCodes = new string[] { };
                if (TopicsList != null && TopicsList.Count > 0)
                {
                    topicCodes = TopicsList.Select(topic => topic.Code).ToArray();
                }
                return topicCodes;
            }
        }

        public string[] GetExportParameters(string topicName)
        {
            if (TopicsList.Count > 0)
            {
                Topic topic = TopicsList.FirstOrDefault(t => String.Compare(t.Code, topicName, StringComparison.OrdinalIgnoreCase) == 0);
                if (topic != null)
                {
                    return topic.Parameters.Export.Select(x => x.Name).ToArray();
                }
            }
            return null;
        }

        public string[] GetImportParameters(string topicName)
        {
            if (TopicsList.Count > 0)
            {
                Topic topic = TopicsList.FirstOrDefault(t => String.Compare(t.Code, topicName, StringComparison.OrdinalIgnoreCase) == 0);
                if (topic != null)
                {
                    return topic.Parameters.Import.Select(x => x.Name).ToArray();
                }
            }
            return null;
        }

        public List<ParameterValue> GetExportParameterValues(string topicName, string parameterName)
        {
            if (TopicsList.Count > 0)
            {
                Topic topic = TopicsList.FirstOrDefault(t => String.Compare(t.Code, topicName, StringComparison.OrdinalIgnoreCase) == 0);
                if (topic != null)
                {
                    Parameter parameter =
                        topic.Parameters.Export.FirstOrDefault(
                            p => String.Compare(p.Name, parameterName, StringComparison.OrdinalIgnoreCase) == 0);
                    if (parameter != null && parameter.ParameterValues != null && parameter.ParameterValues.MultiSelect == 0)
                    {
                        return parameter.ParameterValues.ParameterValueList;
                    }
                }
            }
            return null;
        }

        public List<ParameterValue> GetImportParameterValues(string topicName, string parameterName)
        {
            if (TopicsList.Count > 0)
            {
                Topic topic = TopicsList.FirstOrDefault(t => String.Compare(t.Code, topicName, StringComparison.OrdinalIgnoreCase) == 0);
                if (topic != null)
                {
                    Parameter parameter =
                        topic.Parameters.Import.FirstOrDefault(
                            p => String.Compare(p.Name, parameterName, StringComparison.OrdinalIgnoreCase) == 0);
                    if (parameter != null && parameter.ParameterValues != null && parameter.ParameterValues.MultiSelect == 0)
                    {
                        return parameter.ParameterValues.ParameterValueList;
                    }
                }
            }
            return null;
        }

        public List<ParameterValue> GetExportParameterMultiSelectValues(string topicName, string parameterName)
        {
            if (TopicsList.Count > 0)
            {
                Topic topic = TopicsList.FirstOrDefault(t => String.Compare(t.Code, topicName, StringComparison.OrdinalIgnoreCase) == 0);
                if (topic != null)
                {
                    Parameter parameter =
                        topic.Parameters.Export.FirstOrDefault(
                            p => String.Compare(p.Name, parameterName, StringComparison.OrdinalIgnoreCase) == 0);
                    if (parameter != null && parameter.ParameterValues != null && parameter.ParameterValues.MultiSelect == 1)
                    {
                        return parameter.ParameterValues.ParameterValueList;
                    }
                }
            }
            return null;
        }

        public List<ParameterValue> GetImportParameterMultiSelectValues(string topicName, string parameterName)
        {
            if (TopicsList.Count > 0)
            {
                Topic topic = TopicsList.FirstOrDefault(t => String.Compare(t.Code, topicName, StringComparison.OrdinalIgnoreCase) == 0);
                if (topic != null)
                {
                    Parameter parameter =
                        topic.Parameters.Import.FirstOrDefault(
                            p => String.Compare(p.Name, parameterName, StringComparison.OrdinalIgnoreCase) == 0);
                    if (parameter != null && parameter.ParameterValues != null && parameter.ParameterValues.MultiSelect == 1)
                    {
                        return parameter.ParameterValues.ParameterValueList;
                    }
                }
            }
            return null;
        }

    }

    public class Topic
    {
        [XmlAttribute("code")]
        public string Code { get; set; }

        [XmlElement("Parameters")]
        public Parameters Parameters { get; set; }
    }

    public class Parameters
    {
        [XmlArray("Export"), XmlArrayItem("Parameter", typeof(Parameter))]
        public List<Parameter> Export { get; set; }

        [XmlArray("Import"), XmlArrayItem("Parameter", typeof(Parameter))]
        public List<Parameter> Import { get; set; }
    }

    public class Parameter
    {
        [XmlAttribute("name")]
        public string Name { get; set; }

        [XmlAttribute("description")]
        public string Description { get; set; }

        [XmlAttribute("type")]
        public string Type { get; set; }

        [XmlElement("ParameterValues")]
        public ParameterValues ParameterValues { get; set; }
    }

    public class ParameterValues
    {
        [XmlAttribute("multiselect")]
        public int MultiSelect { get; set; }

        [XmlElement("ParameterValue")]
        public List<ParameterValue> ParameterValueList { get; set; }
    }

    public class ParameterValue
    {
        [XmlAttribute("value")]
        public string Value { get; set; }

        [XmlAttribute("description")]
        public string Description { get; set; }
    }
}
