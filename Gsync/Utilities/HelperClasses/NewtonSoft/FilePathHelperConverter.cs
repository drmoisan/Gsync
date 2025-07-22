using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Gsync.Utilities.Interfaces;
using Gsync.Utilities.Extensions;


namespace Gsync.Utilities.HelperClasses.NewtonSoft
{
    public class FilePathHelperConverter : JsonConverter<FilePathHelper>
    {
        protected FilePathHelperConverter() { }

        public FilePathHelperConverter(IFileSystemFolderPaths fileSystemFolders)
        {
            FileSystemFolders = fileSystemFolders;
        }

        protected IFileSystemFolderPaths _fileSystemFolders;
        internal virtual IFileSystemFolderPaths FileSystemFolders { get => _fileSystemFolders; set => _fileSystemFolders = value; }

        public override FilePathHelper ReadJson(JsonReader reader, Type objectType, FilePathHelper existingValue, bool hasExistingValue, JsonSerializer serializer)
        {
            var filePathHelper = new FilePathHelper();
            var info = ReadToDictionary(reader);
            var folderPath = ExtractFolderPath(info) ?? "";
            var fileName = ExtractFileName(info) ?? "";
            return new FilePathHelper(fileName, folderPath);
        }

        internal string ExtractFolderPath(Dictionary<string, string> info)
        {
            if (!info.TryGetValue("SpecialFolderName", out string folderName)) { return null; }
            if (FileSystemFolders.SpecialFolders.TryGetValue(folderName, out string folderPath))
            { 
                if (info.TryGetValue("RelativePath", out string relativePath))
                {
                    return Path.Combine(folderPath, relativePath);
                }
                else
                {
                    return folderPath; 
                }
            }
            else
            { 
                return null; 
            }
        }

        internal string ExtractFolderPath(string specialFolderName, string relativePath)
        {
            if (specialFolderName != "None" && FileSystemFolders.SpecialFolders.TryGetValue(specialFolderName, out string folderPath))
            {
                if (relativePath.IsNullOrEmpty())
                {
                    return folderPath;
                }
                else
                {
                    return Path.Combine(folderPath, relativePath);
                }
            }

            return relativePath;
        }

        internal string ExtractFileName(Dictionary<string, string> info)
        {
            if (info.TryGetValue("FileName", out string fileName))
            {
                return fileName;
            }
            else
            {
                throw new JsonReaderException("FileName property not found in JSON object.");
            }
        }

        internal Dictionary<string, string> ReadToDictionary(JsonReader reader)
        {
            var info = new Dictionary<string, string>();
            reader.Read();
            while (reader.TokenType != JsonToken.EndObject)
            {                
                string key = ReadPropertyName(reader);
                reader.Read();
                string value = ReadPropertyValue(reader);
                reader.Read();

                info.Add(key, value);
            }
            return info;
        }

        internal string ReadPropertyName(JsonReader reader)
        {
            if (reader.TokenType != JsonToken.PropertyName)
            {
                string message = $"{GetErrorMessage(reader)}. Reader found a token of type " +
                    $"{reader.TokenType} when it was expecting a token of type {JsonToken.PropertyName}."; 
                throw new JsonReaderException(message);
            }

            try
            {
                return (reader.Value as string).ThrowIfNull();
            }
            catch (ArgumentNullException e)
            {
                string message = $"{GetErrorMessage(reader)}. Reader found a token with a property " +
                    $"name that was null or empty.";
                throw new JsonReaderException(message, e);
            }
            
        }

        private static string GetErrorMessage(JsonReader reader)
        {
            var message = $"{nameof(FilePathHelperConverter)}.{nameof(ReadJson)} encountered a problem";
            if (reader is JsonTextReader)
            {
                var textReader = reader as JsonTextReader;
                message += $" on line {textReader.LineNumber} ({reader.Path}).";
            }
            return message;
        }

        internal string ReadPropertyValue(JsonReader reader)
        {
            //if (reader.TokenType != JsonToken.String)            
            Enum[] tokenTypes = [JsonToken.String, JsonToken.Null];
            if (!tokenTypes.Contains(reader.TokenType))
            {
                var message = $"{GetErrorMessage(reader)}. Reader found a token of type {reader.TokenType} " +
                    $"when it was expecting a token of type {JsonToken.String}.";
                throw new JsonReaderException(message);
            }
            return reader.Value as string ?? "";
        }

        public (string Name, string RelativePath) GetSerializablePath(string folderPath)
        {
            string name = "";
            string relativePath = "";
            KeyValuePair<string, string> match = default;

            var matchingFolders = FileSystemFolders.SpecialFolders
                .Where(x => folderPath.Contains(x.Value));

            if (matchingFolders.Count() == 0) 
            { 
                match = new KeyValuePair<string,string>("Not Found", ""); 
            }
            else if (matchingFolders.Count() == 1) 
            {
                match = matchingFolders.First();
            }
            else if (matchingFolders.Count() > 1)
            {
                var max = matchingFolders.Max(x => x.Value.Length);
                match = matchingFolders.First(x => x.Value.Length == max);
            }
               
            name = match.Key;
            relativePath = match.Value.Length > 0 ? folderPath.Replace(match.Value, ""): folderPath;
                        
            return (name, relativePath);
        }

        public override void WriteJson(JsonWriter writer, FilePathHelper value, JsonSerializer serializer)
        {
            var (name, relativePath) = GetSerializablePath(value.FolderPath);

            writer.WriteStartObject();

            writer.WritePropertyName("FileName");
            writer.WriteValue(value.FileName);

            writer.WritePropertyName("RelativePath");
            writer.WriteValue(relativePath);

            writer.WritePropertyName("SpecialFolderName");
            writer.WriteValue(name);

            writer.WriteEndObject();
        }
    }
}
