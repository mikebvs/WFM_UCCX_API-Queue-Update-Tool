using System;
using System.Xml.Serialization;
using System.Collections.Generic;

namespace UCCX_API
{
	[XmlRoot(ElementName = "primarySupervisor")]
	public class PrimarySupervisor
	{
		[XmlElement(ElementName = "refURL")]
		public string RefURL { get; set; }
		[XmlAttribute(AttributeName = "name")]
		public string Name { get; set; }
	}

	[XmlRoot(ElementName = "secondrySupervisor")]
	public class SecondrySupervisor
	{
		[XmlElement(ElementName = "refURL")]
		public string RefURL { get; set; }
		[XmlAttribute(AttributeName = "name")]
		public string Name { get; set; }
	}

	[XmlRoot(ElementName = "secondarySupervisors")]
	public class SecondarySupervisors
	{
		[XmlElement(ElementName = "secondrySupervisor")]
		public List<SecondrySupervisor> SecondrySupervisor { get; set; }
	}

	[XmlRoot(ElementName = "team")]
	public class TeamData
	{
		[XmlElement(ElementName = "self")]
		public string Self { get; set; }
		[XmlElement(ElementName = "teamId")]
		public string TeamId { get; set; }
		[XmlElement(ElementName = "teamname")]
		public string Teamname { get; set; }
		[XmlElement(ElementName = "primarySupervisor")]
		public PrimarySupervisor PrimarySupervisor { get; set; }
		[XmlElement(ElementName = "secondarySupervisors")]
		public SecondarySupervisors SecondarySupervisors { get; set; }
	}

	[XmlRoot(ElementName = "teams")]
	public class Teams
	{
		[XmlElement(ElementName = "team")]
		public List<TeamData> Team { get; set; }
	}

}
