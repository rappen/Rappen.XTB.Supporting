using Rappen.XTB.Helpers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using XrmToolBox.AppCode.AppInsights;
using XrmToolBox.Extensibility;
using XrmToolBox.ToolLibrary.AppCode;

namespace Rappen.XTB
{
    public partial class Supporting : Form
    {
        internal static readonly Uri GeneralSettingsURL = new Uri("https://rappen.github.io/Tools/");

        private static Installation installation;
        private static Tool tool;
        private static Supporters supporters;
        private static ToolSettings settings;
        private static SupportableTool supportabletool;
        private static Random random = new Random();
        private static AppInsights appinsights;

        private readonly Stopwatch sw = new Stopwatch();
        private readonly Stopwatch swInfo = new Stopwatch();

        #region Static Public Methods

        public static void ShowIf(PluginControlBase plugin, bool manual, bool reload, AppInsights appins)
        {
            var toolname = plugin?.ToolName;
            appinsights = appins;
            try
            {
                VerifySettings(toolname, reload);
                VerifyTool(toolname, reload);
                if (supportabletool?.Enabled != true)
                {
                    if (manual)
                    {
                        var url = tool.GetUrlGeneral();
                        appinsights?.WriteEvent($"Supporting-{tool.Acronym}-General");
                        UrlUtils.OpenUrl(url);
                    }
                    return;
                }
                VerifySupporters(toolname, reload);
                CheckUnSubmittedSupporters();
                if (!manual)
                {
                    if (supporters.Any(s => s.Type != SupportType.None && s.Type != SupportType.Never))
                    {   // I have supportings!
                        return;
                    }
                    else if (!supportabletool.ShowAutomatically)
                    {   // Centerally stopping showing automatically
                        return;
                    }
                    else if (tool.Support.Type == SupportType.Never)
                    {   // You will never want to support this tool
                        return;
                    }
                    else if (tool.FirstRunDate.AddMinutes(settings.ShowMinutesAfterToolInstall) > DateTime.Now)
                    {   // Installed it too soon
                        return;
                    }
                    else if (tool.VersionRunDate > tool.FirstRunDate &&
                        tool.VersionRunDate.AddMinutes(settings.ShowMinutesAfterToolNewVersion) > DateTime.Now)
                    {   // Installed this version too soon
                        return;
                    }
                    else if (tool.Support.AutoDisplayDate.AddMinutes(settings.ShowMinutesAfterSupportingShown) > DateTime.Now)
                    {   // Seen this form to soon
                        return;
                    }
                    else if (tool.Support.AutoDisplayCount >= settings.ShowAutoRepeatTimes)
                    {   // Seen this too many times
                        return;
                    }
                    else if (tool.Support.SubmittedDate.AddMinutes(settings.ShowMinutesAfterSubmitting) > DateTime.Now)
                    {   // Submitted too soon for JR to handle it
                        return;
                    }
                    else if (supportabletool.ShowAutoPercentChance < 1 ||
                        supportabletool.ShowAutoPercentChance <= random.Next(1, 100))
                    {
                        return;
                    }
                }
                appinsights?.WriteEvent($"Supporting-{tool.Acronym}-Open-{(manual ? "Manual" : "Auto")}");
                new Supporting(manual).ShowDialog(plugin);
                if (!manual)
                {
                    tool.Support.AutoDisplayDate = DateTime.Now;
                    tool.Support.AutoDisplayCount++;
                }
                installation.Save();
            }
            catch (Exception ex)
            {
                plugin.LogError($"ToolSupporting error:\n{ex}");
            }
        }

        public static bool IsEnabled(PluginControlBase plugin)
        {
            var toolname = plugin?.ToolName;
            VerifySettings(toolname);
            return supportabletool?.Enabled == true;
        }

        public static SupportType IsSupporting(PluginControlBase plugin)
        {
            var toolname = plugin?.ToolName;
            VerifySettings(toolname);
            VerifyTool(toolname);
            VerifySupporters(toolname);
            if (supporters.Any(s => s.Type == SupportType.Company))
            {
                return SupportType.Company;
            }
            if (supporters.Any(s => s.Type == SupportType.Personal))
            {
                return SupportType.Personal;
            }
            if (supporters.Any(s => s.Type == SupportType.Contribute))
            {
                return SupportType.Contribute;
            }
            if (supporters.Any(s => s.Type == SupportType.Already))
            {
                return SupportType.Already;
            }
            if (supporters.Any(s => s.Type == SupportType.Never))
            {
                return SupportType.Never;
            }
            return SupportType.None;
        }

        #endregion Static Public Methods

        #region Static Private Methods

        private static void VerifySettings(string toolname, bool reload = false)
        {
            if (reload || settings == null || supportabletool == null)
            {
                settings = ToolSettings.Get();
                supportabletool = settings[toolname];
                //settings.Save();    // this is only to get a correct format of the tool settings file
            }
        }

        private static void VerifyTool(string toolname, bool reload = false)
        {
            if (reload || installation == null || tool == null)
            {
                supporters = null;
                installation = Installation.Load(settings);
                tool = installation[toolname];
                var version = Assembly.GetExecutingAssembly().GetName().Version;
                if (tool.version != version)
                {
                    tool.version = version;
                    tool.VersionRunDate = DateTime.Now;
                    installation.Save();
                }
            }
        }

        private static void VerifySupporters(string toolname, bool reload = false)
        {
            if (reload || supporters == null)
            {
                supporters = Supporters.DownloadMy(installation.Id, toolname, supportabletool.ContributionCounts);
            }
        }

        private static void CheckUnSubmittedSupporters()
        {
            if (tool.Support.Type != SupportType.None &&
                tool.Support.Type != SupportType.Never &&
                tool.Support.SubmittedDate > DateTime.MinValue &&
                tool.Support.SubmittedDate.AddDays(settings.ResetUnfinalizedSupportingAfterDays) < DateTime.Now &&
                supporters?.Any(s => s.Type == tool.Support.Type) == false)
            {
                tool.Support.Type = SupportType.None;
                tool.Support.SubmittedDate = DateTime.MinValue;
                installation.Save();
            }
        }

        #endregion Static Private Methods

        #region Private Constructors

        private Supporting(bool manual)
        {
            InitializeComponent();
            lblHeader.Text = tool.Name;
            panInfo.Left = 32;
            panInfo.Top = 25;
            SetRandomPositions();
            SetStoredValues(manual);
        }

        #endregion Private Constructors

        #region Private Methods

        private void SetRandomPositions()
        {
            if (settings?.BMACLinkPositionRandom == true)
            {
                if (random.Next(100) < 50)
                {
                    picBuyMeACoffee.Left = btnWhatWhy.Left - 2;
                }
                else
                {
                    picBuyMeACoffee.Left = btnInfo.Left - 2;
                }
            }
            if (settings?.CloseLinkPositionRandom == true)
            {
                var left = random.Next(0, 100);
                var top = random.Next(0, 100);
                if (left < 40) left = settings.CloseLinkHorizFromOrigMin;
                else if (left > 60) left = settings.CloseLinkHorizFromOrigMax;
                else left = (settings.CloseLinkHorizFromOrigMin + settings.CloseLinkHorizFromOrigMax) / 2;
                if (top < 40) top = settings.CloseLinkVertiFromOrigMin;
                else if (top > 60) top = settings.CloseLinkVertiFromOrigMax;
                else top = (settings.CloseLinkVertiFromOrigMin + settings.CloseLinkVertiFromOrigMax) / 2;
                linkClose.Left += left;
                linkClose.Top += top;
            }
        }

        private void SetStoredValues(bool manual = false)
        {
            txtCompanyName.Text = installation.CompanyName;
            txtCompanyEmail.Text = installation.CompanyEmail;
            chkCompanySendInvoice.Checked = installation.SendInvoice;
            txtCompanyCountry.Text = installation.CompanyCountry;
            txtPersonalFirst.Text = installation.PersonalFirstName;
            txtPersonalLast.Text = installation.PersonalLastName;
            txtPersonalEmail.Text = installation.PersonalEmail;
            txtPersonalCountry.Text = installation.PersonalCountry;
            chkPersonalContactMe.Checked = installation.PersonalContactMe;
            if (tool.Support.Type == SupportType.Personal)
            {
                rbPersonal.Checked = true;
            }
            else
            {
                rbCompany.Checked = true;
            }
            if (manual)
            {
                toolTip1.SetToolTip(linkClose, "Close this window.");
            }
            SetAlreadyLink();
            ResetAllColors();
        }

        private void SetAlreadyLink()
        {
            linkStatus.Tag = null;
            bool clickable = true;
            var supporter = supporters?.OrderByDescending(s => s.Date).FirstOrDefault(s => s.Type != SupportType.None);
            switch (supporter?.Type)
            {
                case SupportType.Company:
                    linkStatus.Text = settings.StatusCompanyText.Replace("{tool}", tool.Name);
                    toolTip1.SetToolTip(linkStatus, settings.StatusCompanyTip.Replace("{tool}", tool.Name));
                    break;

                case SupportType.Personal:
                    linkStatus.Text = settings.StatusPersonalText.Replace("{tool}", tool.Name);
                    toolTip1.SetToolTip(linkStatus, settings.StatusPersonalTip.Replace("{tool}", tool.Name));
                    break;

                case SupportType.Contribute:
                    linkStatus.Text = settings.StatusContributeText.Replace("{tool}", tool.Name);
                    toolTip1.SetToolTip(linkStatus, settings.StatusContributeTip.Replace("{tool}", tool.Name));
                    break;

                case SupportType.Already:
                    linkStatus.Text = settings.StatusAlreadyText.Replace("{tool}", tool.Name);
                    toolTip1.SetToolTip(linkStatus, settings.StatusAlreadyTip.Replace("{tool}", tool.Name));
                    break;

                case SupportType.Never:
                    linkStatus.Text = settings.StatusNeverText.Replace("{tool}", tool.Name);
                    toolTip1.SetToolTip(linkStatus, settings.StatusNeverTip.Replace("{tool}", tool.Name));
                    clickable = false;
                    break;

                default:
                    switch (tool.Support.Type)
                    {
                        case SupportType.Company:
                        case SupportType.Personal:
                        case SupportType.Contribute:
                        case SupportType.Already:
                            linkStatus.Text = settings.StatusPendingText.Replace("{tool}", tool.Name);
                            toolTip1.SetToolTip(linkStatus, settings.StatusPendingTip.Replace("{tool}", tool.Name));
                            clickable = false;
                            break;

                        case SupportType.Never:
                            linkStatus.Text = settings.StatusNeverText.Replace("{tool}", tool.Name);
                            toolTip1.SetToolTip(linkStatus, settings.StatusNeverTip.Replace("{tool}", tool.Name));
                            clickable = false;
                            break;

                        default:
                            linkStatus.Text = settings.StatusDefaultText.Replace("{tool}", tool.Name);
                            toolTip1.SetToolTip(linkStatus, settings.StatusDefaultTip.Replace("{tool}", tool.Name));
                            linkStatus.Tag = SupportType.Already;
                            break;
                    }
                    break;
            }
            if (clickable)
            {
                linkStatus.LinkArea = new LinkArea(0, linkStatus.Text.Length - 1);
            }
            else
            {
                linkStatus.LinkArea = new LinkArea(0, 0);
            }
        }

        private void ResetAllColors()
        {
            panBgBlue.BackColor = settings.clrBackground;
            panInfoBg.BackColor = settings.clrBackground;
            helpText.BackColor = settings.clrBackground;
            rbCompany.ForeColor = rbCompany.Checked ? settings.clrTxtFgNormal : settings.clrTxtFgDimmed;
            rbPersonal.ForeColor = rbPersonal.Checked ? settings.clrTxtFgNormal : settings.clrTxtFgDimmed;
            rbContribute.ForeColor = rbContribute.Checked ? settings.clrTxtFgNormal : settings.clrTxtFgDimmed;
            txtCompanyName.BackColor = settings.clrFldBgNormal;
            txtCompanyEmail.BackColor = settings.clrFldBgNormal;
            txtCompanyCountry.BackColor = settings.clrFldBgNormal;
            txtPersonalFirst.BackColor = settings.clrFldBgNormal;
            txtPersonalLast.BackColor = settings.clrFldBgNormal;
            txtPersonalEmail.BackColor = settings.clrFldBgNormal;
            txtPersonalCountry.BackColor = settings.clrFldBgNormal;
            linkStatus.ForeColor = settings.clrTxtFgDimmed;
            linkClose.LinkColor = settings.clrTxtFgDimmed;
        }

        private void SettingSupportType(SupportType type)
        {
            switch (type)
            {
                case SupportType.Already:
                    if (CallingWebForm(tool.GetUrlAlready(), type))
                    {
                        DialogResult = DialogResult.Yes;
                    }
                    break;
            }
        }

        private bool CallingWebForm(string url, SupportType type)
        {
            if (!string.IsNullOrEmpty(url))
            {
                if (MessageBoxEx.Show(this, settings.ConfirmDirecting, "Supporting", MessageBoxButtons.OK, MessageBoxIcon.Asterisk) == DialogResult.OK)
                {
                    tool.Support.Type = type;
                    tool.Support.SubmittedDate = DateTime.Now;
                    appinsights?.WriteEvent($"Supporting-{tool.Acronym}-{tool.Support.Type}");
                    Process.Start(url);
                    return true;
                }
            }
            return false;
        }

        #endregion Private Methods

        #region Private Event Methods

        private void Supporting_Shown(object sender, EventArgs e)
        {
            sw.Restart();
        }

        private void Supporting_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (DialogResult != DialogResult.Yes && DialogResult != DialogResult.OK && DialogResult != DialogResult.Retry)
            {
                e.Cancel = true;
            }
            else
            {
                linkClose.Focus();
                sw.Stop();
                appinsights?.WriteEvent($"Supporting-{tool.Acronym}-Close-{DialogResult}", duration: sw.ElapsedMilliseconds);
            }
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            ctrl_Validating();
            var type = rbPersonal.Checked ? SupportType.Personal :
                rbContribute.Checked ? SupportType.Contribute : SupportType.Company;
            var url = rbPersonal.Checked ? tool.GetUrlPersonal(false) :
                rbContribute.Checked ? tool.GetUrlPersonal(true) : tool.GetUrlCorp();
            if (CallingWebForm(url, type))
            {
                DialogResult = DialogResult.Yes;
            }
        }

        private void ctrl_Validating(object sender = null, System.ComponentModel.CancelEventArgs e = null)
        {
            if (sender == null || sender == txtCompanyName)
            {
                installation.CompanyName = txtCompanyName.Text.Trim().Length >= 3 ? txtCompanyName.Text.Trim() : "";
                txtCompanyName.BackColor = string.IsNullOrEmpty(installation.CompanyName) ? settings.clrFldBgInvalid : settings.clrFldBgNormal;
            }
            if (sender == null || sender == txtCompanyEmail)
            {
                try
                {
                    installation.CompanyEmail = new MailAddress(txtCompanyEmail.Text).Address.Trim();
                }
                catch { }
                txtCompanyEmail.BackColor = string.IsNullOrEmpty(installation.CompanyEmail) ? settings.clrFldBgInvalid : settings.clrFldBgNormal;
            }
            if (sender == null || sender == txtCompanyCountry)
            {
                installation.CompanyCountry = txtCompanyCountry.Text.Trim().Length >= 2 ? txtCompanyCountry.Text.Trim() : "";
                txtCompanyCountry.BackColor = string.IsNullOrEmpty(installation.CompanyCountry) ? settings.clrFldBgInvalid : settings.clrFldBgNormal;
            }
            if (sender == null || sender == chkCompanySendInvoice)
            {
                installation.SendInvoice = chkCompanySendInvoice.Checked;
            }
            if (sender == null || sender == txtPersonalFirst)
            {
                installation.PersonalFirstName = txtPersonalFirst.Text.Trim().Length >= 1 ? txtPersonalFirst.Text.Trim() : "";
                txtPersonalFirst.BackColor = string.IsNullOrEmpty(installation.PersonalFirstName) ? settings.clrFldBgInvalid : settings.clrFldBgNormal;
            }
            if (sender == null || sender == txtPersonalLast)
            {
                installation.PersonalLastName = txtPersonalLast.Text.Trim().Length >= 2 ? txtPersonalLast.Text.Trim() : "";
                txtPersonalLast.BackColor = string.IsNullOrEmpty(installation.PersonalLastName) ? settings.clrFldBgInvalid : settings.clrFldBgNormal;
            }
            if (sender == null || sender == txtPersonalEmail)
            {
                try
                {
                    installation.PersonalEmail = new MailAddress(txtPersonalEmail.Text).Address.Trim();
                }
                catch { }
                txtPersonalEmail.BackColor = string.IsNullOrEmpty(installation.PersonalEmail) ? settings.clrFldBgInvalid : settings.clrFldBgNormal;
            }
            if (sender == null || sender == txtPersonalCountry)
            {
                installation.PersonalCountry = txtPersonalCountry.Text.Trim().Length >= 2 ? txtPersonalCountry.Text.Trim() : "";
                txtPersonalCountry.BackColor = string.IsNullOrEmpty(installation.PersonalCountry) ? settings.clrFldBgInvalid : settings.clrFldBgNormal;
            }
            if (sender == null || sender == chkPersonalContactMe)
            {
                installation.PersonalContactMe = chkPersonalContactMe.Checked;
            }
        }

        private void linkClose_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            DialogResult = DialogResult.Retry;
        }

        private void linkStatus_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (linkStatus.Tag is SupportType type)
            {
                SettingSupportType(type);
            }
            else
            {
                MessageBoxEx.Show("Thanks! ❤️", "Supporting");
            }
        }

        private void rbType_CheckedChanged(object sender, EventArgs e)
        {
            SuspendLayout();
            rbCompany.ForeColor = rbCompany.Checked ? settings.clrTxtFgNormal : settings.clrTxtFgDimmed;
            rbPersonal.ForeColor = rbPersonal.Checked ? settings.clrTxtFgNormal : settings.clrTxtFgDimmed;
            rbContribute.ForeColor = rbContribute.Checked ? settings.clrTxtFgNormal : settings.clrTxtFgDimmed;
            panPersonal.Left = panCorp.Left;
            panPersonal.Top = panCorp.Top;
            lblPersonalIntro.Text = rbContribute.Checked ? "I will contribute with my experience and knowledge!" : "I will monetarily support this tool!";
            panPersonal.Visible = rbPersonal.Checked || rbContribute.Checked;
            panCorp.Visible = !panPersonal.Visible;
            btnSubmit.ImageIndex = rbPersonal.Checked ? 1 : rbContribute.Checked ? 2 : 0;
            ResumeLayout();
        }

        private void helpText_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            TopMost = false;
            UrlUtils.OpenUrl(e.LinkText);
        }

        private void btnWhatWhy_Click(object sender, EventArgs e)
        {
            var visible = panInfo.Tag != sender || !panInfo.Visible;
            panInfo.Tag = sender;
            if (visible)
            {
                helpTitle.Text = settings.HelpWhyTitle;
                helpText.Text = string.Empty;
                helpText.Text = settings.HelpWhyText.Replace("\r\n", "\n").Replace("\n", "\r\n");
            }
            panInfo.Visible = visible;
        }

        private void btnInfo_Click(object sender, EventArgs e)
        {
            var visible = panInfo.Tag != sender || !panInfo.Visible;
            panInfo.Tag = sender;
            if (visible)
            {
                helpTitle.Text = settings.HelpInfoTitle;
                helpText.Text = string.Empty;
                helpText.Text = settings.HelpInfoText.Replace("\r\n", "\n").Replace("\n", "\r\n");
            }
            panInfo.Visible = visible;
        }

        private void btnInfoClose_Click(object sender, EventArgs e)
        {
            panInfo.Visible = false;
        }

        private void panInfo_VisibleChanged(object sender, EventArgs e)
        {
            if (panInfo.Visible)
            {
                swInfo.Restart();
            }
            else
            {
                swInfo.Stop();
                appinsights?.WriteEvent($"Supporting-{tool.Acronym}-{(panInfo.Tag == btnWhatWhy ? "Why" : "Info")}", duration: swInfo.ElapsedMilliseconds);
            }
        }

        private void tsmiLater_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Retry;
        }

        private void tsmiAlready_Click(object sender, EventArgs e)
        {
            SettingSupportType(SupportType.Already);
        }

        private void tsmiNever_Click(object sender, EventArgs e)
        {
            if (MessageBoxEx.Show(this, $"Are you really sure that you don't like {tool.Name}?", "Supporting", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) != DialogResult.Yes)
            {
                if (tool.Support.Type == SupportType.Never)
                {
                    tool.Support.Type = SupportType.None;
                }
                return;
            }
            appinsights?.WriteEvent($"Supporting-{tool.Acronym}-Never");
            tool.Support.Type = SupportType.Never;
            DialogResult = DialogResult.Yes;
        }

        private void tsmiReset_Click(object sender, EventArgs e)
        {
            if (MessageBoxEx.Show(this, "Reset will remove all locally stored data regarding supporting.\nAnything submitted to Jonas will not be removed. If that is needed, please contact me directly.\n\nConfirm reset with Yes/No.", "Supporting", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
            {
                return;
            }
            var toolname = tool.Name;
            installation.Remove();
            VerifyTool(toolname, true);
            VerifySupporters(toolname, true);
            installation.Save();
            SetStoredValues();
        }

        private void lbl_MouseEnter(object sender, EventArgs e)
        {
            if (sender is RadioButton rb)
            {
                rb.ForeColor = settings.clrTxtFgNormal;
            }
            else if (sender is LinkLabel link)
            {
                link.LinkColor = settings.clrTxtFgNormal;
            }
            else if (sender is Label lbl)
            {
                lbl.ForeColor = settings.clrTxtFgNormal;
            }
        }

        private void lbl_MouseLeave(object sender, EventArgs e)
        {
            if (sender is RadioButton rb)
            {
                if (!rb.Checked)
                {
                    rb.ForeColor = settings.clrTxtFgDimmed;
                }
            }
            else if (sender is LinkLabel link)
            {
                link.LinkColor = settings.clrTxtFgDimmed;
            }
            else if (sender is Label lbl)
            {
                lbl.ForeColor = settings.clrTxtFgDimmed;
            }
        }

        private void picBuyMeACoffee_Click(object sender, EventArgs e)
        {
            if (UrlUtils.OpenUrl(sender))
            {
                appinsights?.WriteEvent($"Supporting-{tool.Acronym}-BuyMeACoffee");
                DialogResult = DialogResult.Retry;
            }
        }

        private void tsmiShowInstallationId_Click(object sender, EventArgs e)
        {
            MessageBoxEx.Show(this, $"The XrmToolBox Installation Id is:\n{installation.Id}", "Supporting", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion Private Event Methods
    }

    public class ToolSettings
    {
        private const string FileName = "Rappen.XTB.Settings.xml";

        public int SettingsVersion = 1;
        public List<SupportableTool> SupportableTools = new List<SupportableTool>();
        public int ShowMinutesAfterToolInstall = int.MaxValue;    // 60
        public int ShowMinutesAfterToolNewVersion = int.MaxValue; // 120
        public int ShowMinutesAfterSupportingShown = int.MaxValue; // 2880m / 48h / 2d
        public int ShowMinutesAfterSubmitting = int.MaxValue; // 2880m / 48h / 2d
        public int ShowAutoPercentChance = 0;   // Moved to each tool
        public int ShowAutoRepeatTimes = -1; // 10
        public int ResetUnfinalizedSupportingAfterDays = int.MaxValue; // 7
        public bool BMACLinkPositionRandom = false;
        public bool CloseLinkPositionRandom = false;
        public int CloseLinkHorizFromOrigMin = -90;
        public int CloseLinkHorizFromOrigMax = 0;
        public int CloseLinkVertiFromOrigMin = -50;
        public int CloseLinkVertiFromOrigMax = 0;

        public string FormIdCorporate = "wpf17273";
        public string FormIdPersonal = "wpf17612";
        public string FormIdContribute = "wpf17677";
        public string FormIdAlready = "wpf17761";

        public string FormUrlCorporate =
            "https://jonasr.app/supporting-prefilled/" +
            "?{formid}_1_first={firstname}" +
            "&{formid}_1_last={lastname}" +
            "&{formid}_3={companycountry}" +
            "&{formid}_4={invoiceemail}" +
            "&{formid}_13={tool}" +
            //"&{formid}_19={size}" +
            //"&{formid}_24={amount}" +
            "&{formid}_27={company}" +
            "&{formid}_37={sendinvoice}" +
            "&{formid}_31={tool}" +
            "&{formid}_32={version}" +
            "&{formid}_33={instid}";

        public string FormUrlSupporting =
            "https://jonasr.app/supporting/personal-prefilled/" +
            "?{formid}_1_first={firstname}" +
            "&{formid}_1_last={lastname}" +
            "&{formid}_3={country}" +
            "&{formid}_4={email}" +
            "&{formid}_52={contactme}" +
            "&{formid}_13={tool}" +
            "&{formid}_31={tool}" +
            "&{formid}_32={version}" +
            "&{formid}_33={instid}";

        public string FormUrlContribute =
            "https://jonasr.app/supporting/contribute-prefilled/" +
            "?{formid}_1_first={firstname}" +
            "&{formid}_1_last={lastname}" +
            "&{formid}_3={country}" +
            "&{formid}_4={email}" +
            "&{formid}_52={contactme}" +
            "&{formid}_13={tool}" +
            "&{formid}_31={tool}" +
            "&{formid}_32={version}" +
            "&{formid}_33={instid}";

        public string FormUrlAlready =
            "https://jonasr.app/supporting/already/" +
            "?{formid}_1_first={firstname}" +
            "&{formid}_1_last={lastname}" +
            "&{formid}_3={country}" +
            "&{formid}_4={email}" +
            "&{formid}_13={tool}" +
            "&{formid}_31={tool}" +
            "&{formid}_32={version}" +
            "&{formid}_33={instid}";

        public string FormUrlGeneral =
            "https://jonasr.app/supporting/" +
            "?{formid}_13={tool}" +
            "&{formid}_31={tool}" +
            "&{formid}_32={version}" +
            "&{formid}_33={instid}";

        public string ColorBg = "FF0042AD";                     // FF0042AD Dark blue
        public string ColorTextFgNormal = "FFFFFF00";           // FFFFFF00 Yellow
        public string ColorTextFgDimmed = "FFD2B48C";           // FFD2B48C Dim yellow
        public string ColorFieldBgNormal = "FF0063FF";          // FF0063FF Light blue
        public string ColorFieldBgInvalid = "FFF06565";         // FFF06565 Dim red

        public Color clrBackground => GetColor(ColorBg, "FF0042AD");
        public Color clrTxtFgNormal => GetColor(ColorTextFgNormal, "FFFFFF00");
        public Color clrTxtFgDimmed => GetColor(ColorTextFgDimmed, "FFD2B48C");
        public Color clrFldBgNormal => GetColor(ColorFieldBgNormal, "FF0063FF");
        public Color clrFldBgInvalid => GetColor(ColorFieldBgInvalid, "FFF06565");

        public string ConfirmDirecting = @"You will now be redirected to the website form
to finish Your flavor of support.
After the form is submitted, Jonas will handle it soon.

NOTE: It has to be submitted during the next step!";

        public string StatusDefaultText = "Click here if\r\nYou are already\r\nsupporting!";
        public string StatusDefaultTip = "If you have already supported\r\n{tool}\r\nin any way - Click here to let me know,\r\nand this popup will not appear again!";
        public string StatusCompanyText = "Your company\r\nare supporting\r\n{tool}!";
        public string StatusCompanyTip = "We know that your company is supporting\r\n{tool}\r\nThank You!";
        public string StatusPersonalText = "You are\r\nsupporting\r\n{tool}!";
        public string StatusPersonalTip = "We know that you are supporting\r\n{tool}\r\nThank You!";
        public string StatusContributeText = "You are\r\ncontributing to\r\n{tool}!";
        public string StatusContributeTip = "We know that you are contributing to\r\n{tool}\r\nThank You!";
        public string StatusAlreadyText = "You have already\r\nsupported\r\n{tool}!";
        public string StatusAlreadyTip = "We know that you have already supported\r\n{tool}\r\nThank You!";
        public string StatusNeverText = "You will never\r\nsupport\r\n{tool}.";
        public string StatusNeverTip = "For some strange reason,\r\nyou will never support\r\n{tool}\r\nThink again? 😉";
        public string StatusPendingText = "You have recently\r\nsupported.\r\nJonas is processing it\r\n(if You finalized it).";
        public string StatusPendingTip = "It may take hours/days to process the support...\r\nJonas will handle it after you have finalized the web form.\r\n\r\nThank You so much! ❤️";

        public string HelpWhyTitle = "Community Tools are Conscienceware.";

        public string HelpWhyText = @"Some in the Power Platform Community are creating tools.
Some contribute to the community with new ideas, find problems, write documentation, and even solve our bugs.
Most in this community are mainly 'consumers' — they are only using open-source tools.
To me, it's very similar to watching TV. Do you pay for channels, Netflix, Amazon Prime, Spotify, etc.?
To be part of the community, but without contributing, you can simply just pay instead.

Especially when you work in a big corporation, exploiting free tools - only to increase your income - you have a responsibility to participate actively in the community - or pay.
It's good to be able to sleep with a good conscience. Right?

There should be a license called ""Conscienceware"".
But technically, it is simply free to use them. That's a fact.

If you say you are not part of the community, that is incorrect—just using these tools makes you a part of it.

You and your company can now more formally support tools rather than just donating via PayPal or 'Buy Me a Coffee.'

Supporting is not just giving money; it means that you or your company know that you have gained in time and improved your quality by using these tools. If you get something and want to give back—support the development and maintenance of the tools.

To read more about my thoughts, click here: https://jonasr.app/helping/

- Jonas Rapp";

        public string HelpInfoTitle = "Technical Information";

        public string HelpInfoText = @"Your entered name, company, country, email, and amount will not be stored in any official system. The information will only be saved in my personal Excel file and in my own Power Platform app, mostly for me, myself and I to learn even more about how the platform could help us. I also do this to ensure that you can get an invoice, and if needed we need to communicate by email.
The email you share with me, only to me, will never be sold to any company. I won't try to sell anything. Period.

You will receive an official receipt immediately and, if needed, an invoice. Supporting can be done with a credit card. Other options like Google Pay will be available depending on your location. Stripe handles the payment.

When you click the big button here, the information you entered here will be included in the form on my website, jonasr.app, and a few hidden info: tool name, version, and your XrmToolBox 'InstallationId' (a random Guid generated the first time you use the toolbox). If you are curious, you can see how to find your ID on this link: https://jonasr.app/xtb-finding-installationid.

Since I would like to be very clear and transparent - we store your XrmToolBox InstallationId on a server to be able to know that this installation is supporting it in some way. There is nothing about your name, the amount or contribution; I am not interested in hacking this info.

The button in the top-right corner opens this info. You can also right-click on it and find more options, especially:
* I have already supported this tool — use this to tell me that you already support this tool in some way so that this popup prompt will not ask you again.
* I will never support this tool — use it if you think it is a bad idea, and you probably won't use the tool again; it won't ask you again.

For questions, contact me at https://jonasr.app/contact.";

        private ToolSettings()
        { }

        public static ToolSettings Get() => new Uri(Supporting.GeneralSettingsURL, FileName).DownloadXml<ToolSettings>() ?? new ToolSettings();

        public SupportableTool this[string name]
        {
            get
            {
                if (!SupportableTools.Any(st => st.Name == name))
                {
                    SupportableTools.Add(new SupportableTool { Name = name });
                }
                return SupportableTools.FirstOrDefault(st => st.Name == name);
            }
        }

        public void Save()
        {
            if (!Directory.Exists(Paths.SettingsPath))
            {
                Directory.CreateDirectory(Paths.SettingsPath);
            }
            string path = Path.Combine(Paths.SettingsPath, FileName);
            XmlSerializerHelper.SerializeToFile(this, path);
        }

        private Color GetColor(string color, string defaultColor)
        {
            int intColor;
            try
            {
                intColor = int.Parse(color, System.Globalization.NumberStyles.HexNumber);
            }
            catch
            {
                intColor = int.Parse(defaultColor, System.Globalization.NumberStyles.HexNumber);
            }
            return Color.FromArgb(intColor);
        }
    }

    public class SupportableTool
    {
        public string Name;
        public bool Enabled = false;
        public bool ShowAutomatically = false;
        public bool ContributionCounts = true;
        public int ShowAutoPercentChance = 0;   // 25 (0-100)
    }

    public class Supporters : List<Supporter>
    {
        private const string FileName = "Rappen.XTB.Supporters.xml";

        public static Supporters DownloadMy(Guid InstallationId, string toolname, bool contributionCounts)
        {
            var result = new Uri(Supporting.GeneralSettingsURL, FileName).DownloadXml<Supporters>() ?? new Supporters();
            result.Where(s =>
                s.InstallationId != InstallationId ||
                s.ToolName != toolname)
                .ToList().ForEach(s => result.Remove(s));
            return result;
        }

        public override string ToString() => $"{Count} Supporters";
    }

    public class Supporter
    {
        public Guid InstallationId;
        public string ToolName;
        public SupportType Type;
        public DateTime Date;

        public override string ToString() => $"{InstallationId} {Type} {ToolName} {Date}";
    }

    public class Installation
    {
        private const string FileName = "Rappen.XTB.xml";
        private int settingsversion = -1;
        internal ToolSettings toolsettings;

        public int SettingsVersion
        {
            get => settingsversion;
            set
            {
                if (settingsversion != -1 && value != settingsversion && Tools?.Count() > 0)
                {
                    Tools.ForEach(s => s.Support.AutoDisplayCount = 0);
                }
                settingsversion = value;
            }
        }

        public Guid Id = Guid.Empty;
        public DateTime FirstRunDate = DateTime.Now;
        public string CompanyName;
        public string CompanyEmail;
        public string CompanyCountry;
        public bool SendInvoice;
        public string PersonalFirstName;
        public string PersonalLastName;
        public string PersonalEmail;
        public string PersonalCountry;
        public bool PersonalContactMe;
        public List<Tool> Tools = new List<Tool>();

        public static Installation Load(ToolSettings settings)
        {
            string path = Path.Combine(Paths.SettingsPath, FileName);
            var result = new Installation();
            if (File.Exists(path))
            {
                try
                {
                    XmlDocument xmlDocument = new XmlDocument();
                    xmlDocument.Load(path);
                    result = (Installation)XmlSerializerHelper.Deserialize(xmlDocument.OuterXml, typeof(Installation));
                }
                catch { }
            }
            result.Initialize(settings);
            result.Tools.ForEach(t => t.Installation = result);
            return result;
        }

        internal void Initialize(ToolSettings settings)
        {
            toolsettings = settings;
            if (Id.Equals(Guid.Empty))
            {
                Id = InstallationInfo.Instance.InstallationId;
            }
            if (settings.SettingsVersion != settingsversion)
            {
                SettingsVersion = settings.SettingsVersion;
            }
        }

        public void Save()
        {
            if (!Directory.Exists(Paths.SettingsPath))
            {
                Directory.CreateDirectory(Paths.SettingsPath);
            }
            var path = Path.Combine(Paths.SettingsPath, FileName);
            XmlSerializerHelper.SerializeToFile(this, path);
        }

        public void Remove()
        {
            var path = Path.Combine(Paths.SettingsPath, FileName);
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }

        public Tool this[string name]
        {
            get
            {
                if (!Tools.Any(t => t.Name == name))
                {
                    Tools.Add(new Tool(this, name));
                }
                return Tools.FirstOrDefault(t => t.Name == name);
            }
        }

        public override string ToString() => $"{CompanyName} {PersonalFirstName} {PersonalLastName} {Tools.Count}".Trim().Replace("  ", " ");
    }

    public class Tool
    {
        private Version _version;
        private string name;

        internal Installation Installation;

        internal Version version
        {
            get => _version;
            set
            {
                if (value != _version)
                {
                    Support.AutoDisplayCount = 0;
                    if (Support?.Type == SupportType.Never)
                    {
                        Support.Type = SupportType.None;
                    }
                }
                _version = value;
            }
        }

        public string Name
        {
            get => name;
            set
            {
                name = value;
                switch (name)
                {
                    // Tools that don't have three upper cases
                    case "FetchXML Builder":
                        Acronym = "FXB";
                        break;

                    case "UML Diagram Generator":
                        Acronym = "UML";
                        break;

                    case "XrmToolBox Integration Tester":
                        Acronym = "XIT";
                        break;

                    case "Portal Entity Permission Manager":
                        Acronym = "EPM";
                        break;

                    case "XRM Tokens Runner":
                        Acronym = "XTR";
                        break;

                    case "Shuffle Builder":
                        Acronym = "ShB";
                        break;

                    case "Shuffle Runner":
                        Acronym = "ShR";
                        break;

                    case "Shuffle Deployer":
                        Acronym = "ShD";
                        break;

                    // Tools that have three upper cases
                    default:
                        var pattern = @"((?<=^|\s)(\w{1})|([A-Z]))";
                        Acronym = string.Join(string.Empty, Regex.Matches(Name, pattern).OfType<Match>().Select(x => x.Value.ToUpper()));
                        break;
                }
            }
        }

        internal string Acronym { get; private set; }

        public string Version
        {
            get => _version.ToString();
            set { _version = new Version(value ?? "0.0.0.0"); }
        }

        public DateTime FirstRunDate = DateTime.Now;
        public DateTime VersionRunDate;
        public Support Support = new Support();

        private Tool()
        { }

        internal Tool(Installation installation, string name)
        {
            Installation = installation;
            Name = name;
        }

        public string GetUrlCorp()
        {
            if (string.IsNullOrEmpty(Installation.CompanyName) ||
                string.IsNullOrEmpty(Installation.CompanyEmail) ||
                string.IsNullOrEmpty(Installation.CompanyCountry))
            {
                return null;
            }
            return GenerateUrl(Installation.toolsettings.FormUrlCorporate, Installation.toolsettings.FormIdCorporate);
        }

        public string GetUrlPersonal(bool contribute)
        {
            if (string.IsNullOrEmpty(Installation.PersonalFirstName) ||
                string.IsNullOrEmpty(Installation.PersonalLastName) ||
                string.IsNullOrEmpty(Installation.PersonalEmail) ||
                string.IsNullOrEmpty(Installation.PersonalCountry))
            {
                return null;
            }
            return GenerateUrl(contribute ? Installation.toolsettings.FormUrlContribute : Installation.toolsettings.FormUrlSupporting, contribute ? Installation.toolsettings.FormIdContribute : Installation.toolsettings.FormIdPersonal);
        }

        public string GetUrlAlready()
        {
            return GenerateUrl(Installation.toolsettings.FormUrlAlready, Installation.toolsettings.FormIdAlready);
        }

        public string GetUrlGeneral()
        {
            return GenerateUrl(Installation.toolsettings.FormUrlGeneral, Installation.toolsettings.FormIdCorporate);
        }

        private string GenerateUrl(string template, string form)
        {
            return template
                .Replace("{formid}", form)
                .Replace("{company}", Installation.CompanyName)
                .Replace("{invoiceemail}", Installation.CompanyEmail)
                .Replace("{companycountry}", Installation.CompanyCountry)
                .Replace("{amount}", Support.Amount)
                .Replace("{size}", Support.UsersCount)
                .Replace("{sendinvoice}", Installation.SendInvoice ? "Send%20me%20an%20invoice" : "")
                .Replace("{firstname}", Installation.PersonalFirstName)
                .Replace("{lastname}", Installation.PersonalLastName)
                .Replace("{email}", Installation.PersonalEmail)
                .Replace("{country}", Installation.PersonalCountry)
                .Replace("{contactme}", Installation.PersonalContactMe ? "Contact%20me%20after%20submitting%20this%20form!" : "")
                .Replace("{tool}", Name)
                .Replace("{version}", version.ToString())
                .Replace("{instid}", Installation.Id.ToString());
        }

        public override string ToString() => $"{Name} {version}";
    }

    public class Support
    {
        public DateTime AutoDisplayDate = DateTime.MinValue;
        public int AutoDisplayCount;
        public DateTime SubmittedDate;
        public SupportType Type = SupportType.None;
        public int UsersIndex;
        public string UsersCount;

        public string Amount
        {
            get
            {
                switch (UsersIndex)
                {
                    case 1: return "X-Small";
                    case 2: return "Small";
                    case 4: return "Large";
                    case 5: return "X-Large";
                    default: return "Medium";
                }
            }
        }

        public override string ToString() => $"{Type}";
    }

    public enum SupportType
    {
        None,
        Personal,
        Company,
        Contribute,
        Already,
        Never
    }
}