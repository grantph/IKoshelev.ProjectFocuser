using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ProjectFocuser
{
    public class CsProject
    {
        public string Name { get; set; }
        public string Path { get; set; }
        public Guid Guid { get; set; }
        public string Data { get; set; }
        public System.Xml.XmlDocument XmlDocument { get; set; }
    }

    public static class FileParser
    {
        /*
		public static Dictionary<string, string> GetProjectNamesToGuidsDict(string solutionPath)
		{
			var text = File.ReadAllText(solutionPath);

			//Raw regex: /Project\(\"\{([0-9ABCDEF-]*)\}\"\) = \"([^"]*)\", \"([^"]*)\", \"\{([0-9ABCDEF-]*)\}\"/g
			var regex = new Regex(@"Project\(\""\{([0-9ABCDEF-]*)\}\""\) = \""([^""]*)\"", \""([^""]*)\"", \""\{([0-9ABCDEF-]*)\}\""");

			var matches = regex.Matches(text);

			var dict = matches.Cast<Match>()
								.Where(x => x.Groups[3].Value.EndsWith(".csproj", StringComparison.OrdinalIgnoreCase))
								.ToDictionary(
										x => x.Groups[2].Value,
										x => x.Groups[4].Value);

			return dict;
		}
		*/

        public static Dictionary<string, CsProject> GetProjects(string solutionPath)
        {
            var text = File.ReadAllText(solutionPath);

            //Raw regex: /Project\(\"\{([0-9ABCDEF-]*)\}\"\) = \"([^"]*)\", \"([^"]*)\", \"\{([0-9ABCDEF-]*)\}\"/g
            var regex = new Regex(@"Project\(\""\{([0-9ABCDEF-]*)\}\""\) = \""([^""]*)\"", \""([^""]*)\"", \""\{([0-9ABCDEF-]*)\}\""");

            var matches = regex.Matches(text);

            var dict = matches.Cast<Match>()
                                .Where(x => x.Groups[3].Value.EndsWith(".csproj", StringComparison.OrdinalIgnoreCase) ||
                                            x.Groups[3].Value.EndsWith(".sqlproj", StringComparison.OrdinalIgnoreCase) ||
                                            x.Groups[3].Value.EndsWith(".vbproj", StringComparison.OrdinalIgnoreCase))
                                .ToDictionary(
                                        x => x.Groups[2].Value,
                                        x => new CsProject()
                                        {
                                            Name = x.Groups[2].Value,
                                            Path = x.Groups[3].Value,
                                            Guid = new Guid(x.Groups[4].Value)
                                        });

            return dict;
        }
    }
}