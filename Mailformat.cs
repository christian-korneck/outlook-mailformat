// Pure text transforms, ported from chrome-mailformat/background.js (apply()).
using System;
using System.Collections.Generic;
using System.Linq;

namespace MailformatAddin
{
    public static class Mailformat
    {
        public const int DefaultWidth = 72;

        // Leading quote prefix, equivalent to /^(>[> ]*)?/.
        static string Prefix(string line)
        {
            if (line.Length == 0 || line[0] != '>') return "";
            int i = 1;
            while (i < line.Length && (line[i] == '>' || line[i] == ' ')) i++;
            return line.Substring(0, i);
        }

        public static string Wrap(string text, int width)
        {
            return string.Join("\n", text.Split('\n').Select(line =>
            {
                string prefix = Prefix(line);
                string[] words = line.Substring(prefix.Length)
                    .Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
                if (words.Length == 0) return line;
                int max = Math.Max(1, width - prefix.Length);
                var output = new List<string>();
                string cur = "";
                foreach (string word in words)
                {
                    if (cur.Length > 0 && cur.Length + 1 + word.Length > max) { output.Add(cur); cur = word; }
                    else cur = cur.Length > 0 ? cur + " " + word : word;
                }
                output.Add(cur);
                return string.Join("\n", output.Select(l => prefix + l));
            }));
        }

        public static string Quote(string text)
        {
            return string.Join("\n", text.Split('\n').Select(l =>
                l == "" ? ">" : l.StartsWith(">") ? ">" + l : "> " + l));
        }

        public static string Unquote(string text)
        {
            return string.Join("\n", text.Split('\n').Select(l =>
                l.StartsWith("> ") ? l.Substring(2) : l.StartsWith(">") ? l.Substring(1) : l));
        }

        // mode: wrap | quote | unquote | both
        public static string Apply(string mode, string text, int width)
        {
            switch (mode)
            {
                case "wrap": return Wrap(text, width);
                case "quote": return Quote(text);
                case "unquote": return Unquote(text);
                case "both": return Quote(Wrap(text, width - 2));
                default: throw new ArgumentException("unknown mode: " + mode);
            }
        }
    }
}
