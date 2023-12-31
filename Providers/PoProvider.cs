﻿using LocalizePo.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace LocalizePo.Providers
{
    public class PoProvider : IPoProvider
    {
        public void ExportToFile<TModel>(IEnumerable<PoObject<TModel>> data, string filePath, Encoding encoding) where TModel : new()
        {
            var text = string.Join(Environment.NewLine + Environment.NewLine, data.Select(d => d.Serialize()));

            File.WriteAllText(filePath, text, encoding);
        }

        public IEnumerable<PoObject<TModel>> ReadFile<TModel>(string filePath) where TModel : new()
        {
            var result = new List<PoObject<TModel>>();
            var line = string.Empty;
            var objectLines = new List<string>();

            if (!File.Exists(filePath))
            {
                throw new Exception($"File {filePath} not found");
            }

            using (var stringReader = new StringReader(File.ReadAllText(filePath)))
            {
                while ((line = stringReader.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        result.Add(new PoObject<TModel>(objectLines));
                        objectLines.Clear();
                    }
                    else
                    {
                        objectLines.Add(line);
                    }
                }

                result.Add(new PoObject<TModel>(objectLines));
            }

            return result;
        }
    }

    public class PoObject<TModel> where TModel : new()
    {
        private List<PoRow> Rows { get; set; }
        public TModel Model { get; private set; }

        public PoObject(List<string> lines)
        {
            var properties = typeof(TModel).GetProperties();
            Rows = lines.Select(l => new PoRow(l, properties.Select(p => p.GetPropertyName()).ToList())).ToList();
            Model = new TModel();

            foreach (var row in Rows)
            {
                var property = properties.FirstOrDefault(p => p.GetPropertyName() == row.PropertyName);

                if (property != null)
                {
                    property.SetValue(Model, row.PropertyValue);
                }
            }
        }

        public void SetModelPropertyText(PropertyInfo property, string value)
        {
            var row = Rows.FirstOrDefault(r => r.PropertyName == property.GetPropertyName());

            if (row == null)
            {
                throw new Exception($"Row with property {property.GetPropertyName()} not found");
            }

            row.PropertyValue = value;
            property.SetValue(Model, value);
        }

        public string Serialize()
        {
            return string.Join(Environment.NewLine, Rows.Select(r => r.Serialize()));
        }
    }

    public class PoRow
    {
        public string PropertyName { get; set; }
        public string PropertyValue { get; set; }

        public PoRow(string line, List<string> propertyNames)
        {
            PropertyName = propertyNames.FirstOrDefault(pn => line.StartsWith(pn)) ?? string.Empty;
            PropertyValue = string.IsNullOrWhiteSpace(PropertyName) ? line : line.Substring(PropertyName.Length);
        }

        public string Serialize()
        {
            return $"{PropertyName}{PropertyValue}";
        }
    }
}
