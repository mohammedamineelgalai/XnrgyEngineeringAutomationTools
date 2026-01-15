// ============================================================================
// ProjectSelectorWindow.xaml.cs
// Project Selector Dialog - Modern XNRGY Theme
// Author: Mohammed Amine Elgalai - XNRGY Climate Systems ULC
// ============================================================================

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace XnrgyEngineeringAutomationTools.Modules.DXFVerifier.Views
{
    /// <summary>
    /// Fenetre de selection de projet avec theme moderne XNRGY
    /// </summary>
    public partial class ProjectSelectorWindow : Window
    {
        private List<ProjectPathInfo> _allProjects;
        private ICollectionView _projectsView;
        
        /// <summary>
        /// Projet selectionne (null si annule)
        /// </summary>
        public ProjectPathInfo SelectedProject { get; private set; }

        public ProjectSelectorWindow(List<ProjectPathInfo> projects)
        {
            InitializeComponent();
            _allProjects = projects ?? new List<ProjectPathInfo>();
            LoadProjects();
        }

        private void LoadProjects()
        {
            // Trier: Projet decroissant, puis Reference, puis Module
            var sortedProjects = _allProjects
                .OrderByDescending(p => p.ProjectNumber)
                .ThenBy(p => p.Reference)
                .ThenBy(p => p.ModuleNumber)
                .ToList();

            ProjectsDataGrid.ItemsSource = sortedProjects;
            _projectsView = CollectionViewSource.GetDefaultView(ProjectsDataGrid.ItemsSource);
            
            UpdateProjectCount();
        }

        private void UpdateProjectCount()
        {
            var count = _projectsView?.Cast<ProjectPathInfo>().Count() ?? 0;
            var projectCount = _projectsView?.Cast<ProjectPathInfo>().Select(p => p.ProjectNumber).Distinct().Count() ?? 0;
            TxtProjectCount.Text = $"{count} modules dans {projectCount} projets";
        }

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_projectsView == null) return;
            
            var searchText = TxtSearch.Text?.Trim().ToLowerInvariant() ?? "";
            
            if (string.IsNullOrEmpty(searchText))
            {
                _projectsView.Filter = null;
            }
            else
            {
                _projectsView.Filter = obj =>
                {
                    if (obj is ProjectPathInfo project)
                    {
                        return project.ProjectNumber.ToLowerInvariant().Contains(searchText) ||
                               project.Reference.ToLowerInvariant().Contains(searchText) ||
                               project.ModuleNumber.ToLowerInvariant().Contains(searchText) ||
                               project.ProjectPath.ToLowerInvariant().Contains(searchText);
                    }
                    return false;
                };
            }
            
            UpdateProjectCount();
        }

        private void ProjectsDataGrid_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            SelectCurrentProject();
        }

        private void BtnSelect_Click(object sender, RoutedEventArgs e)
        {
            SelectCurrentProject();
        }

        private void SelectCurrentProject()
        {
            if (ProjectsDataGrid.SelectedItem is ProjectPathInfo selected)
            {
                SelectedProject = selected;
                DialogResult = true;
                Close();
            }
            else
            {
                MessageBox.Show("Veuillez selectionner un projet dans la liste.", 
                               "Selection requise", 
                               MessageBoxButton.OK, 
                               MessageBoxImage.Warning);
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            SelectedProject = null;
            DialogResult = false;
            Close();
        }
    }
}
