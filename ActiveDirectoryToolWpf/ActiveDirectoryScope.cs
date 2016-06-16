using System;
using System.Collections.Generic;
using PrimitiveExtensions;

namespace ActiveDirectoryToolWpf
{
    public class ActiveDirectoryScope : IEquatable<ActiveDirectoryScope>
    {
        private const string DomainComponentPrefix = "DC=";
        private const char Comma = ',';
        private const char Period = '.';
        private const string LdapProtocolPrefix = "LDAP://";

        internal ActiveDirectoryScope()
        {
            Children = new List<ActiveDirectoryScope>();
        }

        public List<ActiveDirectoryScope> Children { get; set; }
        internal string Name { get; set; }
        internal string Path { get; set; }
        internal string Context => Path.Remove(LdapProtocolPrefix);

        internal string Domain => Path.SubstringAtIndexOfOrdinal(DomainComponentPrefix)
            .Remove(DomainComponentPrefix)
            .Replace(Comma, Period);

        public bool Equals(ActiveDirectoryScope other)
        {
            return Name == other.Name;
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
                else if (level == organizationalUnit.Split[lastLevel])
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