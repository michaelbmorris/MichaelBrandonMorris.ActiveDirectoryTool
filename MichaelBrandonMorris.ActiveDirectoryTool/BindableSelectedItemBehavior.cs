using System.Windows;
using System.Windows.Controls;
using System.Windows.Interactivity;

namespace MichaelBrandonMorris.ActiveDirectoryTool
{
    /// <summary>
    ///     Class BindableSelectedItemBehavior.
    /// </summary>
    /// <seealso cref="Behavior{T}" />
    /// <seealso cref="TreeView" />
    /// TODO Edit XML Comment Template for BindableSelectedItemBehavior
    public class BindableSelectedItemBehavior : Behavior<TreeView>
    {
        /// <summary>
        ///     The selected item property
        /// </summary>
        /// TODO Edit XML Comment Template for SelectedItemProperty
        public static readonly DependencyProperty SelectedItemProperty =
            DependencyProperty.Register(
                "SelectedItem",
                typeof(object),
                typeof(BindableSelectedItemBehavior),
                new UIPropertyMetadata(null, OnSelectedItemChanged));

        /// <summary>
        ///     Gets or sets the selected item.
        /// </summary>
        /// <value>The selected item.</value>
        /// TODO Edit XML Comment Template for SelectedItem
        public object SelectedItem
        {
            get => GetValue(SelectedItemProperty);
            set => SetValue(SelectedItemProperty, value);
        }

        /// <summary>
        ///     Called when [attached].
        /// </summary>
        /// TODO Edit XML Comment Template for OnAttached
        protected override void OnAttached()
        {
            base.OnAttached();

            AssociatedObject.SelectedItemChanged +=
                OnTreeViewSelectedItemChanged;
        }

        /// <summary>
        ///     Called when [detaching].
        /// </summary>
        /// TODO Edit XML Comment Template for OnDetaching
        protected override void OnDetaching()
        {
            base.OnDetaching();

            if (AssociatedObject != null)
            {
                AssociatedObject.SelectedItemChanged -=
                    OnTreeViewSelectedItemChanged;
            }
        }

        /// <summary>
        ///     Handles the <see cref="E:SelectedItemChanged" /> event.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">
        ///     The <see cref="DependencyPropertyChangedEventArgs" />
        ///     instance containing the event data.
        /// </param>
        /// TODO Edit XML Comment Template for OnSelectedItemChanged
        private static void OnSelectedItemChanged(
            DependencyObject sender,
            DependencyPropertyChangedEventArgs e)
        {
            var item = e.NewValue as TreeViewItem;
            item?.SetValue(TreeViewItem.IsSelectedProperty, true);
        }

        /// <summary>
        ///     Called when [TreeView selected item changed].
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The e.</param>
        /// TODO Edit XML Comment Template for OnTreeViewSelectedItemChanged
        private void OnTreeViewSelectedItemChanged(
            object sender,
            RoutedPropertyChangedEventArgs<object> e)
        {
            SelectedItem = e.NewValue;
        }
    }
}