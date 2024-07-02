using System;
using System.Drawing;
using System.Windows.Forms;

namespace BIMSearch
{
    public class SearchForm : Form
    {
        private TextBox textBox;
        private ComboBox levelComboBox;
        private Button searchButton;
        private Button createSectionBoxButton;
        private Button cancelButton;

        public string SearchTerm { get; private set; }
        public string SelectedLevel { get; private set; }

        public event EventHandler SearchClicked;
        public event EventHandler CreateSectionBoxClicked;

        public SearchForm(string[] levels)
        {
            InitializeComponent(levels);
        }

        private void InitializeComponent(string[] levels)
        {
            this.textBox = new TextBox();
            this.levelComboBox = new ComboBox();
            this.searchButton = new Button();
            this.createSectionBoxButton = new Button();
            this.cancelButton = new Button();

            this.SuspendLayout();

            // textBox
            this.textBox.Location = new Point(12, 12);
            this.textBox.Name = "textBox";
            this.textBox.Size = new Size(260, 20);
            this.textBox.TabIndex = 0;

            // levelComboBox
            this.levelComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            this.levelComboBox.Location = new Point(12, 38);
            this.levelComboBox.Name = "levelComboBox";
            this.levelComboBox.Size = new Size(260, 21);
            this.levelComboBox.TabIndex = 1;
            this.levelComboBox.Items.Add("All Levels");
            foreach (var level in levels)
            {
                this.levelComboBox.Items.Add(level);
            }
            this.levelComboBox.SelectedIndex = 0;

            // searchButton
            this.searchButton.Location = new Point(12, 65);
            this.searchButton.Name = "searchButton";
            this.searchButton.Size = new Size(75, 23);
            this.searchButton.TabIndex = 2;
            this.searchButton.Text = "Search";
            this.searchButton.UseVisualStyleBackColor = true;
            this.searchButton.Click += new EventHandler(this.SearchButton_Click);

            // createSectionBoxButton
            this.createSectionBoxButton.Location = new Point(93, 65);
            this.createSectionBoxButton.Name = "createSectionBoxButton";
            this.createSectionBoxButton.Size = new Size(120, 23);
            this.createSectionBoxButton.TabIndex = 3;
            this.createSectionBoxButton.Text = "Create Section Box";
            this.createSectionBoxButton.UseVisualStyleBackColor = true;
            this.createSectionBoxButton.Click += new EventHandler(this.CreateSectionBoxButton_Click);
            this.createSectionBoxButton.Enabled = false; // initially disabled

            // cancelButton
            this.cancelButton.Location = new Point(219, 65);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new Size(75, 23);
            this.cancelButton.TabIndex = 4;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new EventHandler(this.CancelButton_Click);

            // SearchForm
            this.ClientSize = new Size(306, 101);
            this.Controls.Add(this.textBox);
            this.Controls.Add(this.levelComboBox);
            this.Controls.Add(this.searchButton);
            this.Controls.Add(this.createSectionBoxButton);
            this.Controls.Add(this.cancelButton);
            this.Name = "SearchForm";
            this.Text = "Search for Elements";

            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void SearchButton_Click(object sender, EventArgs e)
        {
            this.SearchTerm = this.textBox.Text;
            this.SelectedLevel = this.levelComboBox.SelectedItem.ToString();
            this.SearchClicked?.Invoke(this, EventArgs.Empty);
        }

        private void CreateSectionBoxButton_Click(object sender, EventArgs e)
        {
            this.CreateSectionBoxClicked?.Invoke(this, EventArgs.Empty);
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void EnableCreateSectionBoxButton(bool enable)
        {
            this.createSectionBoxButton.Enabled = enable;
        }
    }
}
