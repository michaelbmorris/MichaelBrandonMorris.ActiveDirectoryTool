using System;
using System.Collections.Generic;

namespace ActiveDirectoryToolWpf
{
    public class ActiveDirectoryScope : IEquatable<ActiveDirectoryScope>
    {
        public List<ActiveDirectoryScope> Children { get; set; }
        internal string Name { get; set; }
        internal string Path { get; set; }

        internal string Domain
        {
            get
            {
                return Path.Substring(Path.IndexOf("DC=", StringComparison.Ordinal))
                        .Replace("DC=", string.Empty)
                        .Replace(",", ".");
            }
        }

        internal ActiveDirectoryScope()
        {
            Children = new List<ActiveDirectoryScope>();
        }

        public bool Equals(ActiveDirectoryScope other)
        {
            return Name.Equals(other.Name);
        }

        public override string ToString()
        {
            return Name;
        }

        internal void AddDirectoryScope(OrganizationalUnit organizationalUnit)
        {
            if (organizationalUnit.Split == null ||
                organizationalUnit.Split.Length < 1)
            {
                throw new ArgumentException(
                    "The organizational units array is null or empty!");
            }
            var parent = this;
            foreach (var level in organizationalUnit.Split)
            {
                var lastLevel = organizationalUnit.Split.Length - 1;
                if (parent.Children.Contains(new ActiveDirectoryScope
                {
                    Name = level
                }))
                {
                    parent = parent.Children.Find(
                        item => item.Name.Equals(level));
                }
                else if (level.Equals(organizationalUnit.Split[lastLevel]))
                {
                    parent.Children.Add(new ActiveDirectoryScope
                    {
                        Name = level,
                        Path = organizationalUnit.Path
                    });
                }
            }
        }
    }
}