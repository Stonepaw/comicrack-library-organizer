using System.ComponentModel;
using System.Resources;
using System.Windows.Forms;
using LibraryOrganizer.ViewModel;

namespace LibraryOrganizer.Controls
{
    [ComplexBindingProperties("DataSource")]
    internal partial class InsertTemplateControl : UserControl
    {
        [Category("Data")]
        [Description("The templates to display")]
        [RefreshProperties(RefreshProperties.Repaint)]
        [AttributeProvider(typeof(IListSource))]
        public object DataSource
        {
            get => templateViewModelBindingSource.DataSource;
            set => templateViewModelBindingSource.DataSource = value;
        }

        public InsertTemplateControl()
        {
            InitializeComponent();
            templateViewModelBindingSource.CurrentChanged += (sender, args) =>
            {
                CurrentTemplateChanged();
            };
        }

        private void CurrentTemplateChanged()
        {
            ITemplateFormatViewModel format = (
                (TemplateViewModel)templateViewModelBindingSource.Current
            ).Format;

            switch (format)
            {
                case NumberInsertTemplateFormatConfig numberFormat:
                    UseNumberFormatControl(numberFormat);
                    break;
                case YesNoInsertTemplateFormatConfig yesNoFormat:
                    UseYesNoFormatControl(yesNoFormat);
                    break;
                default:
                    ClearFormatControl();
                    break;
            }
        }

        private void ClearFormatControl()
        {
            tableLayoutPanel1.SuspendLayout();
            panel1.SuspendLayout();
            panel1.Controls.Clear();

            tableLayoutPanel1.ColumnStyles[3] = new ColumnStyle(SizeType.Absolute, 0);
            panel1.ResumeLayout();
            tableLayoutPanel1.ResumeLayout();
        }

        private void UseNumberFormatControl(NumberInsertTemplateFormatConfig format)
        {
            Control existing = panel1.Controls.Count > 0 ? panel1.Controls[0] : null;

            if (existing is NumberTemplateFormatControl)
            {
                // TODO: update the model
                return;
            }

            tableLayoutPanel1.SuspendLayout();
            panel1.SuspendLayout();
            panel1.Controls.Clear();

            tableLayoutPanel1.ColumnStyles[3] = new ColumnStyle(SizeType.AutoSize);
            NumberTemplateFormatControl numberTemplateFormatControl =
                new NumberTemplateFormatControl();
            numberTemplateFormatControl.Dock = DockStyle.Fill;
            panel1.Controls.Add(numberTemplateFormatControl);

            panel1.ResumeLayout();
            tableLayoutPanel1.ResumeLayout();
        }

        private void UseYesNoFormatControl(YesNoInsertTemplateFormatConfig format)
        {
            Control existing = panel1.Controls.Count > 0 ? panel1.Controls[0] : null;

            if (existing is YesNoTemplateFormatControl)
            {
                // TODO: update the model
                return;
            }

            tableLayoutPanel1.SuspendLayout();
            panel1.SuspendLayout();
            panel1.Controls.Clear();

            tableLayoutPanel1.ColumnStyles[3] = new ColumnStyle(SizeType.AutoSize);
            YesNoTemplateFormatControl yesNoTemplateFormatControl =
                new YesNoTemplateFormatControl();
            yesNoTemplateFormatControl.Dock = DockStyle.Fill;
            panel1.Controls.Add(yesNoTemplateFormatControl);

            panel1.ResumeLayout();
            tableLayoutPanel1.ResumeLayout();
        }
    }
}
