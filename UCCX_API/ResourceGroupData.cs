using System;
using System.Xml.Serialization;
using System.Collections.Generic;

namespace UCCX_API
{
	[XmlRoot(ElementName = "resourceGroup")]
	public class ResourceGroupData
	{
		[XmlElement(ElementName = "self")]
		public string Self { get; set; }
		[XmlElement(ElementName = "id")]
		public string Id { get; set; }
		[XmlElement(ElementName = "name")]
		public string Name { get; set; }
	}

	[XmlRoot(ElementName = "resourceGroups")]
	public class ResourceGroups
	{
		[XmlElement(ElementName = "resourceGroup")]
		public List<ResourceGroupData> ResourceGroup { get; set; }
	}

}
