using System.Runtime.Serialization;

namespace DCE.PCF.Plugins.Modals
{
    [DataContract]
    public class PolyLookupItem
    {
        [DataMember(Name = "id")]
        public string Id { get; set; }
        [DataMember(Name = "name")]
        public string Name { get; set; }
        [DataMember(Name = "etn")]
        public string Etn { get; set; }
    }
}
