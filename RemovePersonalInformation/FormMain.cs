using System;
using System.IO;
using System.Windows.Forms;

namespace RemovePersonalInformation
{
    public partial class FormMain : Form
    {
        #region Public Methods

        public FormMain()
        {
            InitializeComponent();
        }

        #endregion

        #region Private Methods

        private void doRemove(string[] files)
        {
            DialogResult dialogResult = showMessage("Office ファイルから個人情報を削除します。", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

            if (dialogResult != DialogResult.OK)
            {
                return;
            }

            string[] completedFiles;
            string[] faultFiles;

            try
            {
                MainEngine.RemovePersonalInformation(files, out completedFiles, out faultFiles);
            }
            catch (Exception exception)
            {
                showErrorMessage(exception.Message);
                return;
            }

            string message;

            if (faultFiles.Length > 0)
            {
                message = "以下のファイルの個人情報は削除できませんでした。\r\n";
                Array.ForEach(faultFiles, n => message += ("\r\n" + Path.GetFileName(n)));
                showErrorMessage(message);
            }

            if (completedFiles.Length > 0)
            {
                message = "以下の Office ファイルから個人情報を削除しました。\r\n";
                Array.ForEach(completedFiles, n => message += ("\r\n" + Path.GetFileName(n)));
                showMessage(message, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void showErrorMessage(string text)
        {
            showMessage(text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private DialogResult showMessage(string text, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            return MessageBox.Show(this, text, Text, buttons, icon);
        }

        #endregion

        // Designer Methods

        private void dragDrop(object sender, DragEventArgs e)
        {
            var dropData = (e.Data.GetData(DataFormats.FileDrop)) as string[];

            if ((dropData == null) || (dropData.Length < 1))
            {
                return;
            }

            doRemove(dropData);
        }

        private void dragEnter(object sender, DragEventArgs e)
        {
            e.Effect
                = (e.Data.GetDataPresent(DataFormats.FileDrop)
                ? DragDropEffects.All
                : DragDropEffects.None);
        }
    }
}
