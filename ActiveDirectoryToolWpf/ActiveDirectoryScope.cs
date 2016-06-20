using System;
using System.Collections.Generic;
using PrimitiveExtensions;

namespace ActiveDirectoryToolWpf
{
    public class ActiveDirectoryScope : IEquatable<ActiveDirectoryScope>
    {
        private const char Comma = ',';
        private const string DomainComponentPrefix = "DC=";
        private const string LdapProtocolPrefix = "LDAP://";
        private const char Period = '.';

        internal ActiveDirectoryScope()
        {
            Children = new List<ActiveDirectoryScope>();
        }

        public List<ActiveDirectoryScope> Children { get; set; }
        internal string Context => Path.Remove(LdapProtocolPrefix);

        internal string Domain
            => Path.SubstringAtIndexOfOrdinal(DomainComponentPrefix)
                .Remove(DomainComponentPrefix)
                .Replace(Comma, Period);

        internal string Name { get; set; }
        internal string Path { get; set; }

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
            var organizationalUnitLevels = organizationalUnit.Split();
            if (organizationalUnitLevels == null ||
                organizationalUnitLevels.Length < 1)
            {
                throw new ArgumentException(
                    "The organizational units array is null or empty!");
            }

            var parent = this;
            var lastLevelIndex = organizationalUnitLevels.Length - 1;
            foreach (var level in organizationalUnitLevels)
            {
                var scope = new ActiveDirectoryScope
                {
                    Name = level
                };
                if (parent.Children.Contains(scope))
                {
                    parent = parent.Children.Find(
                        item => item.Name.Equals(level));
                }
                else if (level == organizationalUnitLevels[lastLevelIndex])
                {
                    scope.Path = organizationalUnit.Path;
                    parent.Children.Add(scope);
                }
            }
        }
    }
}