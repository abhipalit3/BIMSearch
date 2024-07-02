using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace BIMSearch
{
    [Plugin("BIMSearchPlugin", "DPR", DisplayName = "BIM Search Plugin")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class BIMSearchPlugin : AddInPlugin
    {
        private SearchForm searchForm;
        private ModelItemCollection searchResults;

        public override int Execute(params string[] parameters)
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            if (doc == null)
            {
                MessageBox.Show("No active document found.");
                return 0;
            }

            string[] levels = GetLevels(doc);
            searchForm = new SearchForm(levels);
            searchForm.SearchClicked += OnSearchClicked;
            searchForm.CreateSectionBoxClicked += OnCreateSectionBoxClicked;

            searchForm.Show();

            return 0;
        }

        private string[] GetLevels(Document doc)
        {
            HashSet<string> levels = new HashSet<string>();

            Search search = new Search();
            search.Selection.SelectAll();

            // Create search condition to find level properties
            SearchCondition condition = SearchCondition.HasPropertyByDisplayName("Item", "Level");

            // Apply search condition
            search.SearchConditions.Add(condition);

            // Execute search
            ModelItemCollection results = search.FindAll(doc, false);

            foreach (ModelItem item in results)
            {
                DataProperty levelProperty = item.PropertyCategories.FindPropertyByDisplayName("Item", "Level");
                if (levelProperty != null)
                {
                    levels.Add(levelProperty.Value.ToDisplayString());
                }
            }

            return levels.ToArray();
        }

        private void OnSearchClicked(object sender, EventArgs e)
        {
            string searchTerm = searchForm.SearchTerm;
            string selectedLevel = searchForm.SelectedLevel;
            PerformSearch(searchTerm, selectedLevel);
        }

        private void OnCreateSectionBoxClicked(object sender, EventArgs e)
        {
            if (searchResults != null && searchResults.Count > 0)
            {
                CreateSectionBox(searchResults);
            }
        }

        private void PerformSearch(string searchTerm, string selectedLevel)
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            if (doc == null)
            {
                MessageBox.Show("No active document found.");
                return;
            }

            string logFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "NavisworksPropertiesLog.txt");

            try
            {
                using (StreamWriter writer = new StreamWriter(logFilePath, false))
                {
                    searchResults = SearchModelItems(doc, searchTerm, selectedLevel, writer);
                    HandleSearchResults(doc, searchResults, searchTerm, logFilePath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
        }

        private ModelItemCollection SearchModelItems(Document doc, string searchTerm, string selectedLevel, StreamWriter writer)
        {
            ModelItemCollection results = new ModelItemCollection();
            ModelItem rootItem = doc.Models.First.RootItem;

            foreach (var item in rootItem.DescendantsAndSelf)
            {
                if (IsMatch(item, searchTerm, selectedLevel, writer))
                {
                    results.Add(item);
                }
            }

            return results;
        }

        private bool IsMatch(ModelItem item, string searchTerm, string selectedLevel, StreamWriter writer)
        {
            bool matchFound = false;
            bool levelMatch = selectedLevel.Equals("All Levels");

            foreach (var propertyCategory in item.PropertyCategories)
            {
                foreach (var property in propertyCategory.Properties)
                {
                    string propertyValue = GetPropertyValueAsString(property);
                    writer.WriteLine($"Item: {item.DisplayName}, Category: {propertyCategory.DisplayName}, Property: {property.DisplayName}, Value: {propertyValue}");

                    if (propertyCategory.DisplayName.Equals("Item", StringComparison.OrdinalIgnoreCase))
                    {
                        if (property.DisplayName.Equals("Name", StringComparison.OrdinalIgnoreCase) &&
                            propertyValue.IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            matchFound = true;
                        }

                        if (property.DisplayName.Equals("Level", StringComparison.OrdinalIgnoreCase) &&
                            propertyValue.Equals(selectedLevel, StringComparison.OrdinalIgnoreCase))
                        {
                            levelMatch = true;
                        }
                    }
                }
            }

            return matchFound && levelMatch;
        }

        private void HandleSearchResults(Document doc, ModelItemCollection results, string searchTerm, string logFilePath)
        {
            if (results.Count == 0)
            {
                MessageBox.Show($"No elements found with the specified name: {searchTerm}. Properties have been logged to {logFilePath}");
                searchForm.EnableCreateSectionBoxButton(false);
            }
            else
            {
                doc.CurrentSelection.CopyFrom(results);
                MessageBox.Show($"{results.Count} elements found with the name containing '{searchTerm}'.");
                searchForm.EnableCreateSectionBoxButton(true);

                FocusOnCurrentSelection(doc);
            }
        }

        private string GetPropertyValueAsString(DataProperty property)
        {
            try
            {
                if (property.Value.IsDisplayString)
                {
                    return property.Value.ToDisplayString();
                }
                else if (property.Value.IsNamedConstant)
                {
                    return property.Value.ToNamedConstant().DisplayName;
                }
                else if (property.Value.IsInt32)
                {
                    return property.Value.ToInt32().ToString();
                }
                else if (property.Value.IsDouble)
                {
                    return property.Value.ToDouble().ToString();
                }
                else if (property.Value.IsBoolean)
                {
                    return property.Value.ToBoolean().ToString();
                }
                else
                {
                    return "Unsupported property type";
                }
            }
            catch (Exception ex)
            {
                return $"Error retrieving value: {ex.Message}";
            }
        }

        private void CreateSectionBox(ModelItemCollection items)
        {
            if (items.Count == 0) return;

            BoundingBox3D boundingBox = items[0].BoundingBox();

            foreach (ModelItem item in items)
            {
                BoundingBox3D itemBox = item.BoundingBox();
                boundingBox = new BoundingBox3D(
                    new Point3D(
                        Math.Min(boundingBox.Min.X, itemBox.Min.X),
                        Math.Min(boundingBox.Min.Y, itemBox.Min.Y),
                        Math.Min(boundingBox.Min.Z, itemBox.Min.Z)),
                    new Point3D(
                        Math.Max(boundingBox.Max.X, itemBox.Max.X),
                        Math.Max(boundingBox.Max.Y, itemBox.Max.Y),
                        Math.Max(boundingBox.Max.Z, itemBox.Max.Z))
                );
            }

            var sectionBox = new
            {
                Type = "ClipPlaneSet",
                Version = 1,
                OrientedBox = new
                {
                    Type = "OrientedBox3D",
                    Version = 1,
                    Box = new[]
                    {
                        new[] { boundingBox.Min.X, boundingBox.Min.Y, boundingBox.Min.Z },
                        new[] { boundingBox.Max.X, boundingBox.Max.Y, boundingBox.Max.Z }
                    },
                    Rotation = new[] { 0, 0, 0 }
                },
                Enabled = true
            };

            string clippingPlanesJson = JsonConvert.SerializeObject(sectionBox, Formatting.Indented);

            using (Transaction transaction = Autodesk.Navisworks.Api.Application.ActiveDocument.BeginTransaction("Create Section Box"))
            {
                Autodesk.Navisworks.Api.Application.ActiveDocument.ActiveView.SetClippingPlanes(clippingPlanesJson);
                transaction.Commit();
            }

            MessageBox.Show("Section box created around the selected elements.");

            FocusOnCurrentSelection(Autodesk.Navisworks.Api.Application.ActiveDocument);
        }

        private void FocusOnCurrentSelection(Document doc)
        {
            doc.ActiveView.FocusOnCurrentSelection();
        }
    }
}
