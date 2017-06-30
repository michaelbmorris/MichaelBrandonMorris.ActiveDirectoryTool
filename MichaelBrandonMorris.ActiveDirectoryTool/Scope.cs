using System;
using System.Collections.Generic;
using MichaelBrandonMorris.Extensions.PrimitiveExtensions;

namespace MichaelBrandonMorris.ActiveDirectoryTool
{
    /// <summary>
    ///     Class Scope.
    /// </summary>
    /// <seealso
    ///     cref="System.IEquatable{MichaelBrandonMorris.ActiveDirectoryTool.Scope}" />
    /// TODO Edit XML Comment Template for Scope
    public class Scope : IEquatable<Scope>
    {
        /// <summary>
        ///     The comma
        /// </summary>
        /// TODO Edit XML Comment Template for Comma
        private const char Comma = ',';

        /// <summary>
        ///     The domain component prefix
        /// </summary>
        /// TODO Edit XML Comment Template for DomainComponentPrefix
        private const string DomainComponentPrefix = "DC=";

        /// <summary>
        ///     The LDAP prefix
        /// </summary>
        /// TODO Edit XML Comment Template for LdapPrefix
        private const string LdapPrefix = "LDAP://";

        /// <summary>
        ///     The period
        /// </summary>
        /// TODO Edit XML Comment Template for Period
        private const char Period = '.';

        /// <summary>
        ///     Initializes a new instance of the <see cref="Scope" />
        ///     class.
        /// </summary>
        /// TODO Edit XML Comment Template for #ctor
        internal Scope()
        {
            Children = new List<Scope>();
        }

        /// <summary>
        ///     Gets or sets the children.
        /// </summary>
        /// <value>The children.</value>
        /// TODO Edit XML Comment Template for Children
        public List<Scope> Children
        {
            get;
            set;
        }

        /// <summary>
        ///     Gets the context.
        /// </summary>
        /// <value>The context.</value>
        /// TODO Edit XML Comment Template for Context
        internal string Context => Path.Remove(LdapPrefix);

        /// <summary>
        ///     Gets the domain.
        /// </summary>
        /// <value>The domain.</value>
        /// TODO Edit XML Comment Template for Domain
        internal string Domain => Path
            .SubstringAtIndexOfOrdinal(DomainComponentPrefix)
            .Remove(DomainComponentPrefix)
            .Replace(Comma, Period);

        /// <summary>
        ///     Gets or sets the name.
        /// </summary>
        /// <value>The name.</value>
        /// TODO Edit XML Comment Template for Name
        internal string Name
        {
            get;
            set;
        }

        /// <summary>
        ///     Gets or sets the path.
        /// </summary>
        /// <value>The path.</value>
        /// TODO Edit XML Comment Template for Path
        internal string Path
        {
            get;
            set;
        }

        /// <summary>
        ///     Indicates whether the current object is equal to
        ///     another object of the same type.
        /// </summary>
        /// <param name="other">An object to compare with this object.</param>
        /// <returns>
        ///     true if the current object is equal to the
        ///     <paramref name="other" /> parameter; otherwise, false.
        /// </returns>
        /// TODO Edit XML Comment Template for Equals
        public bool Equals(Scope other)
        {
            return other != null && Name == other.Name;
        }

        /// <summary>
        ///     Returns a <see cref="System.String" /> that represents
        ///     this instance.
        /// </summary>
        /// <returns>
        ///     A <see cref="System.String" /> that represents
        ///     this instance.
        /// </returns>
        /// TODO Edit XML Comment Template for ToString
        public override string ToString()
        {
            return Name;
        }

        /// <summary>
        ///     Adds the directory scope.
        /// </summary>
        /// <param name="organizationalUnit">The organizational unit.</param>
        /// <exception cref="ArgumentException">
        ///     The organizational
        ///     units array is null or empty!
        /// </exception>
        /// TODO Edit XML Comment Template for AddDirectoryScope
        internal void AddDirectoryScope(OrganizationalUnit organizationalUnit)
        {
            var organizationalUnitLevels = organizationalUnit.Split();
            if (organizationalUnitLevels == null
                || organizationalUnitLevels.Length < 1)
            {
                throw new ArgumentException(
                    "The organizational units array is null or empty!");
            }

            var parent = this;
            var lastLevelIndex = organizationalUnitLevels.Length - 1;
            foreach (var level in organizationalUnitLevels)
            {
                var scope = new Scope
                {
                    Name = level
                };
                if (parent.Children.Contains(scope))
                {
                    parent =
                        parent.Children.Find(item => item.Name.Equals(level));
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