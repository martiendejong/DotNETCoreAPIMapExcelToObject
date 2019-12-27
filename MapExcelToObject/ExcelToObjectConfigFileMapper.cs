using AutoMapper;
using IniParser;
using IniParser.Exceptions;
using IniParser.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelToObject
{
    public class ExcelToObjectConfigFileMapper<MappedObjectType> : ExcelToObjectMapper<MappedObjectType>
    {
        private static IMapper CreateMapper(string mappingFolder = "") => new Mapper(CreateMapperConfiguration(mappingFolder));


        public ExcelToObjectConfigFileMapper() : base(CreateMapper())
        {
        }

        public ExcelToObjectConfigFileMapper(string mappingFolder) : base(CreateMapper(mappingFolder))
        {
        }

        // Automapper configuration
        private static MapperConfiguration CreateMapperConfiguration(string mappingFolder)
        {
            var config = new MapperConfiguration(cfg =>
            {
                if(!string.IsNullOrEmpty(mappingFolder))
                {
                    mappingFolder = mappingFolder + "\\";
                }
                string configFile = $@"{mappingFolder}{typeof(MappedObjectType).Name}.mapping";
                ConfigurableMapping<MappedObjectType> mapping = LoadMapping(configFile);

                cfg.AddProfiles(new Profile[]
                {
                    mapping
                });
            });

            return config;
        }

        private static ConfigurableMapping<MappedObjectType> LoadMapping(string configFile)
        {
            IniData configFileData;
            try
            {
                FileIniDataParser parser = new FileIniDataParser();
                configFileData = parser.ReadFile(configFile);
            }
            catch (ParsingException parsingException)
            {
                throw new ExcelToObjectException($@"Error parsing mapping file {configFile}: {parsingException.Message}");
            }

            SectionData section = GetSection(configFileData);
            IList<MappedField> fields = GetFields(section);

            ConfigurableMapping<MappedObjectType> mapping = new ConfigurableMapping<MappedObjectType>(section.SectionName, fields);
            return mapping;
        }

        private static IList<MappedField> GetFields(SectionData section)
        {
            IList<MappedField> fields = new List<MappedField>();
            var pairsEnumerator = section.Keys.GetEnumerator();
            KeyData pair;
            while (pairsEnumerator.MoveNext())
            {
                pair = pairsEnumerator.Current;
                fields.Add(new MappedField(pair.Value, pair.KeyName));
            }

            return fields;
        }

        private static SectionData GetSection(IniData configFileData)
        {
            SectionData section;
            var sectionsEnumerator = configFileData.Sections.GetEnumerator();
            if (sectionsEnumerator.MoveNext())
            {
                section = sectionsEnumerator.Current;
            }
            else
            {
                throw new ExcelToObjectException($@"Mapping file should contain a line ""[sheetname]"" ie [Sheet1]");
            }

            return section;
        }
    }
}
