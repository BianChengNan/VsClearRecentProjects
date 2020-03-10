using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace VsClearRecentProjects
{
    public partial class ClearRecentProjectsForm : Form
    {
        public ClearRecentProjectsForm()
        {
            InitializeComponent();
            allCheckBoxs = new CheckBox[]
            {
                checkBox_vs2005,
                checkBox_vs2008,
                checkBox_vs2010,
                checkBox_vs2012,
                checkBox_vs2013,
                checkBox_vs2015,
                checkBox_vs2017,
                checkBox_vs2019,
                checkBox_vsNext,
            };

            vsVersionMapping = new Dictionary<CheckBox, Version>
            {
                { checkBox_vs2005, new Version(8, 0)  },
                { checkBox_vs2008, new Version(9, 0)  },
                { checkBox_vs2010, new Version(10, 0) },
                { checkBox_vs2012, new Version(11, 0) },
                { checkBox_vs2013, new Version(12, 0) },
                { checkBox_vs2015, new Version(14, 0) },
                { checkBox_vs2017, new Version(15, 0) },
                { checkBox_vs2019, new Version(16, 0) },
                { checkBox_vsNext, new Version(17, 0) },
            };

            checkBox_All.Checked = true;
        }

        private void CheckBox_vs_CheckedChanged(object sender, EventArgs e)
        {
            if (sender == null)
            {
                return;
            }

            if (sender == checkBox_All)
            {
                foreach (var checkBox in allCheckBoxs)
                {
                    checkBox.CheckedChanged -= CheckBox_vs_CheckedChanged;
                    checkBox.Checked = checkBox_All.Checked;
                    checkBox.CheckedChanged += CheckBox_vs_CheckedChanged;
                }
            }
            else
            {
                int checkedCount = 0;
                foreach (var checkBox in allCheckBoxs)
                {
                    if (checkBox.Checked)
                    {
                        ++checkedCount;
                    }
                }

                var checkState = CheckState.Unchecked;
                if (checkedCount == 0)
                {
                    checkState = CheckState.Unchecked;
                }
                else if (checkedCount == allCheckBoxs.Length)
                {
                    checkState = CheckState.Checked;
                }
                else
                {
                    checkState = CheckState.Indeterminate;
                }

                checkBox_All.CheckedChanged -= CheckBox_vs_CheckedChanged;
                checkBox_All.CheckState = checkState;
                checkBox_All.CheckedChanged += CheckBox_vs_CheckedChanged;
            }
        }

        private void Button_clear_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (var vsCheckBox in allCheckBoxs)
                {
                    if (vsCheckBox.Checked)
                    {
                        Version vsVersion;
                        if (vsVersionMapping.TryGetValue(vsCheckBox, out vsVersion))
                        {
                            ClearMRU(vsVersion);
                        }
                        else
                        {
                            throw new ArgumentException("not supported vs version");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Remove Key failed with Exception:\n {0}", ex.ToString()));
            }
        }

        private void ClearMRU(Version vsVersion)
        {
            if (vsVersion.Major < (int)VsVersion.Vs2017)
            {
                RemoveFromRegistry(vsVersion);
            }
            else
            {
                RemoveFromFile(vsVersion);
            }
        }

        private void RemoveFromRegistry(Version vsVersion)
        {
            var subKeyPath = string.Format("\\Software\\Microsoft\\VisualStudio\\{0}.{1}\\ProjectMRUList", vsVersion.Major, vsVersion.Minor);
            var subKey = Registry.CurrentUser.OpenSubKey(subKeyPath);
            if (null != subKey)
            {
                Registry.CurrentUser.DeleteSubKey(subKeyPath);
            }
        }

        private void RemoveFromFile(Version vsVersion)
        {
            // reference https://www.pstips.net/navigating-the-file-system.html
            var appDataLocalPath = System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var vsAppDataLocalPath = Path.Combine(appDataLocalPath, "Microsoft\\VisualStudio\\");

            var subFolderPrefix = Path.Combine(vsAppDataLocalPath, string.Format("{0}.{1}", vsVersion.Major, vsVersion.Minor));

            var allMathedFiles = Directory.EnumerateFiles(vsAppDataLocalPath, "ApplicationPrivateSettings.xml", SearchOption.AllDirectories).Where(fullPath =>
            {
                return fullPath.StartsWith(subFolderPrefix);
            }).ToArray();

            if (allMathedFiles.Length > 1)
            {
                string tipMessage = "Multi ApplicationPrivateSettings.xml found, Remove All?\n";
                foreach (var file in allMathedFiles)
                {
                    tipMessage += file + "\n";
                }
                var result = MessageBox.Show(tipMessage, "Warning", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                if (result != DialogResult.Yes)
                {
                    return;
                }
            }

            foreach (var fullPath in allMathedFiles)
            {
                try
                {
                    var backupFileFullPath = fullPath + "-" + DateTime.Now.ToString("yyyy-MM-dd-HH.mm.ss.fff") + ".bak";
                    File.Move(fullPath, backupFileFullPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private CheckBox[] allCheckBoxs;

        private Dictionary<CheckBox, Version> vsVersionMapping;

        enum VsVersion
        {
            Vs2005 = 8,
            Vs2008 = 9,
            Vs2010 = 10,
            Vs2012 = 11,
            Vs2013 = 12,
            Vs2015 = 14,
            Vs2017 = 15,
            Vs2019 = 16,
        }

        private void button_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
