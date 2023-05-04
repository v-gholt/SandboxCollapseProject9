using Community.VisualStudio.Toolkit;
using EnvDTE;
using EnvDTE100;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using SandboxCollapseProject9;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SandboxCollapseProjects9
{
    [Command(PackageIds.ContextMenuCommand)]
    internal sealed class MyCommand : BaseCommand<MyCommand>
    {
        //access DTE
        private static DTE2 _dte;
        private static DTE2 Dte
        {
            get
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                if (_dte == null)
                    _dte = (DTE2)ServiceProvider.GlobalProvider.GetService(typeof(DTE));
                return _dte;
            }
        }

        //get Solution Explorer Window
        private static UIHierarchy s_solutionExplorer;

        private static UIHierarchy SolutionExplorer => s_solutionExplorer ??= Dte.ToolWindows.SolutionExplorer;

        //get selected item in solution explorer
        private static UIHierarchyItem GetSelectedUIHierarchy()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (SolutionExplorer.SelectedItems is object[] array && array.Length == 1)
            {
                return array[0] as UIHierarchyItem;
            }
            return null;
        }

        //recursive method to iterate over children's children to collapse each item in the hierarchy
        void IterateOverChildren(UIHierarchyItem item)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            UIHierarchyItems children = item.UIHierarchyItems;
            if (item != null && children != null)
            {
                foreach (UIHierarchyItem child in children)
                {
                    child.UIHierarchyItems.Expanded = false;
                    IterateOverChildren(child);
                }
            }
        }

        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            UIHierarchyItem selectedUIHierarchy = GetSelectedUIHierarchy();
            UIHierarchyItems items = selectedUIHierarchy.UIHierarchyItems;
            if (selectedUIHierarchy != null && items != null)
            {
                foreach (UIHierarchyItem item in items)
                {
                    IterateOverChildren(item);
                }

                selectedUIHierarchy.UIHierarchyItems.Expanded = false;

            }
        }

    }
}
