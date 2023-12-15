using System.ComponentModel;

namespace LocalizePo.Entities
{
    public class LocalizationData
    {
        [DisplayName("#. Key:\t")]
        public string Key { get; set; }

        [DisplayName("#. SourceLocation:\t")]
        public string SourceLocation { get; set; }

        [DisplayName("msgctxt ")]
        public string Msgctxt { get; set; }

        [DisplayName("msgid ")]
        public string OriginalText { get; set; }

        [DisplayName("msgstr ")]
        public string LocalizedText { get; set; }
    }
}
