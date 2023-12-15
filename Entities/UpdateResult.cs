using System.ComponentModel;

namespace LocalizePo.Entities
{
    public class UpdateResult : LocalizationData
    {
        [DisplayName("Successfully processed")]
        public bool IsSuccessfullyProcessed { get; set; }

        public string Message { get; set; }
    }
}
