namespace FilesEditor.Constants
{
    public class UserErrorMessages
    {
        public const string UnableToOpenFile = "The selected file could not be opened. Make sure that the file is not already open and that it is in a compatible format.";
        public const string UnableToUpdateFile = "The file {0} could not be updated. Make sure that the file is not already open.";
        public const string UnableToDeleteFile = "The tool needs to delete the file '{0}' but this is already in use or protected. Please make sure it is deleted.";
        public const string UnableToDeleteFolder = "The tool needs to delete the folder '{0}' but this is already in use or protected. Please make sure it is deleted.";
        public const string UnableToCreateFolder = "The tool needs to create the file '{0}' but there was an error during the attempt of creating it";

        public const string MissingWorksheet = "The worksheet '{0}' is missing";
        public const string MissingValue = "The required value '{0}' is missing";
        public const string MissingHeader = "The file '{0}' does not containt the expected header '{1}' inside the worksheet '{2}'";
        public const string InvalidValue = "Invalid value '{0}'";
    }
}