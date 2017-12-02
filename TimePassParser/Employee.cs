using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TimePassParser {
	class Employee {
		public string Name { get; set; }
		public int TotalDays { get; set; }
		public TimeSpan TotalHours { get; set; }
		public int TotalPasses { get; set; }
		public Dictionary<string, DayInfo> Days { get; set; }

		public Employee(string name) {
			Name = name;
			TotalDays = 0;
			TotalHours = new TimeSpan();
			TotalPasses = 0;
			Days = new Dictionary<string, DayInfo>();
		}

		public class DayInfo {
			public string PassDate { get; set; }
			public TimeSpan WorkingTime { get; set; }
			public int Passes { get; set; }
			public string FirstEnterTime { get; set; }
			public string LastExitTime { get; set; }

			public DayInfo(string passDate) {
				PassDate = passDate;
				WorkingTime = new TimeSpan();
				Passes = 0;
				FirstEnterTime = string.Empty;
				LastExitTime = string.Empty;
			}
		}
	}
}
