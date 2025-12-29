using Microsoft.Toolkit.Uwp.Notifications;
using Rappen.XTB.Helpers;
using Rappen.XTB.Helpers.RappXTB;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Mail;
using System.Reflection;
using System.Windows.Forms;
using XrmToolBox.Extensibility;

namespace Rappen.XTB
{
    public partial class Supporting : Form
    {
        private static Installation installation;
        private static Tool tool;
        private static Supporters supporters;
        private static SupportableTool supportabletool;
        private static ToastableTool toastabletool;
        private static Random random = new Random();
        private static AppInsights appinsights;
        private static bool isshowing;
        private readonly Stopwatch sw = new Stopwatch();
        private readonly Stopwatch swInfo = new Stopwatch();

        #region Static Public Methods

        public static void ShowIf(RappPluginControlBase plugin, ShowItFrom from, bool manual, bool reload, SupportType? type = null, bool sync = false)
        {
            if (plugin == null)
            {
                return;
            }

            if (sync)
            {
                ShowIfInternal(plugin, from, manual, reload, type);
                return;
            }
            try
            {
                if (plugin.IsHandleCreated)
                {
                    plugin.BeginInvoke((Action)(() => ShowIfInternal(plugin, from, manual, reload, type)));
                }
                else
                {
                    plugin.HandleCreated += (s, e) =>
                        plugin.BeginInvoke((Action)(() => ShowIfInternal(plugin, from, manual, reload, type)));
                }
            }
            catch (Exception ex)
            {
                // Best-effort logging; avoids surfacing exceptions from fire-and-forget
                try
                { plugin.LogError($"Supporting.ShowIfFireAndForget failed:\n{ex}"); }
                catch { }
            }
        }

        public static bool IsEnabled(PluginControlBase plugin)
        {
            var toolname = plugin?.ToolName;
            VerifySettings(toolname);
            return supportabletool?.Enabled == true;
        }

        public static bool IsPending(PluginControlBase plugin)
        {
            var toolname = plugin?.ToolName;
            VerifySettings(toolname);
            VerifyTool(toolname);
            VerifySupporters(toolname);
            if (tool.Support.Type == SupportType.None)
            {
                return false;
            }

            if (tool.Support.Type == SupportType.Never)
            {
                return false;
            }

            if (tool.Support.SubmittedDate == DateTime.MinValue)
            {
                return false;
            }

            if (tool.Support.SubmittedDate.AddMinutes(Settings.Instance.ShowMinutesAfterSubmitting) < DateTime.Now)
            {
                return false;
            }

            return true;
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

        public static bool IsMonetarySupporting(PluginControlBase plugin)
        {
            var supporttype = IsSupporting(plugin);
            return supporttype == SupportType.Company ||
                   supporttype == SupportType.Personal;
        }

        /// <summary>
        /// Handles the activation of a toast notification and performs the appropriate action based on the provided
        /// arguments.
        /// </summary>
        /// <remarks>It *must* have an method in your tools:
        /// public override void HandleToastActivation(ToastNotificationActivatedEventArgsCompat args)
        /// </remarks>
        /// <param name="plugin">The plugin instance that provides context for the operation.</param>
        /// <param name="args">The arguments associated with the toast notification activation.</param>
        /// <param name="ai">The application insights instance used for telemetry and logging.</param>
        /// <returns><see langword="true"/> if the toast activation was successfully handled and an action was performed;
        /// otherwise, <see langword="false"/>.</returns>
        public static bool HandleToastActivation(RappPluginControlBase plugin, ToastNotificationActivatedEventArgsCompat args, AppInsights ai)
        {
            var toastArgs = ToastArguments.Parse(args.Argument);
            if (!toastArgs.TryGetValue("action", out var action) ||
                !toastArgs.TryGetValue("sender", out var sender) ||
                sender != "supporting")
            {
                return false;
            }
            switch (action)
            {
                case "corporate":
                    //ShowIf(plugin, ShowItFrom.ToastCall, true, false, ai, SupportType.Company);
                    VerifySettings(plugin.ToolName);
                    VerifyTool(plugin.ToolName);
                    VerifySupporters(plugin.ToolName);
                    OpenWebForm(tool.GetUrlCorp(false), SupportType.Company);
                    return true;

                case "personal":
                    //ShowIf(plugin, ShowItFrom.ToastCall, true, false, ai, SupportType.Personal);
                    //return true;
                    VerifySettings(plugin.ToolName);
                    VerifyTool(plugin.ToolName);
                    VerifySupporters(plugin.ToolName);
                    OpenWebForm(tool.GetUrlPersonal(false, false), SupportType.Personal);
                    return true;

                case "default":
                    ShowIf(plugin, ShowItFrom.Button, true, false);
                    return true;

                default:
                    return false;
            }
        }

        #endregion Static Public Methods

        #region Static Private Methods

        private static void ShowIfInternal(RappPluginControlBase plugin, ShowItFrom from, bool manual, bool reload, SupportType? type = null)
        {
            var toolname = plugin?.ToolName;
            appinsights = plugin.AppInsights;
            try
            {
                VerifySettings(toolname, reload);
                VerifyTool(toolname, reload);
                if (from == ShowItFrom.Execute)
                {
                    tool.ExecuteCount++;
                    _ = installation.SaveAsync();
                }
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
                if (!manual && supporters.Any(s => s.Type != SupportType.None && s.Type != SupportType.Never))
                {   // I am already supporting!
                    return;
                }
                else if (tool.Support.Type == SupportType.Never)
                {   // You will never want to support this tool
                    return;
                }
                if (ShowIfPopup(plugin, from, manual, type))
                {
                    return;
                }
                if (ShowIfToast(plugin, from))
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                plugin.LogError($"ToolSupporting error:\n{ex}");
            }
        }

        private static bool ShowIfPopup(PluginControlBase plugin, ShowItFrom from, bool manual, SupportType? type)
        {
            if (from == ShowItFrom.Execute)
            {
                return false;
            }
            if (!manual)
            {
                if (!supportabletool.ShowAutomatically)
                {   // Centerally stopping showing automatically
                    return false;
                }
                else if (tool.FirstRunDate.AddMinutes(Settings.Instance.ShowMinutesAfterToolInstall) > DateTime.Now)
                {   // Installed it too soon
                    return false;
                }
                else if (tool.VersionRunDate > tool.FirstRunDate &&
                    tool.VersionRunDate.AddMinutes(Settings.Instance.ShowMinutesAfterToolNewVersion) > DateTime.Now)
                {   // Installed this version too soon
                    return false;
                }
                else if (tool.Support.AutoDisplayDate.AddMinutes(Settings.Instance.ShowMinutesAfterSupportingShown) > DateTime.Now)
                {   // Seen this form to soon
                    return false;
                }
                else if (tool.Support.AutoDisplayCount >= Settings.Instance.ShowAutoRepeatTimes)
                {   // Seen this too many times
                    return false;
                }
                else if (tool.Support.SubmittedDate.AddMinutes(Settings.Instance.ShowMinutesAfterSubmitting) > DateTime.Now)
                {   // Submitted too soon for JR to handle it
                    return false;
                }
                else if (supportabletool.ShowAutoPercentChance < 1 ||
                    supportabletool.ShowAutoPercentChance < random.Next(1, 31))
                {
                    return false;
                }
            }
            if (isshowing)
            {
                return false;
            }
            isshowing = true;
            appinsights?.WriteEvent($"Supporting-{tool.Acronym}-Open-{(manual ? "Manual" : "Auto")}");
            new Supporting(manual, type).ShowDialog(plugin);
            isshowing = false;
            if (!manual)
            {
                tool.Support.AutoDisplayDate = DateTime.Now;
                tool.Support.AutoDisplayCount++;
            }
            _ = installation.SaveAsync();
            return true;
        }

        private static bool ShowIfToast(PluginControlBase plugin, ShowItFrom from)
        {
            if (toastabletool?.Enabled != true)
            {
                return false;
            }
            if (tool.Support.ToastedDate.AddMinutes(toastabletool.MinutesBetweenToasts) > DateTime.Now)
            {   // Don't do it too often
                return false;
            }
            switch (from)
            {
                case ShowItFrom.Open:
                    if (toastabletool.OpenPercentChance < random.Next(1, 101))
                    {   // Random didn't want to toast
                        return false;
                    }
                    break;

                case ShowItFrom.Execute:
                    if (tool.ExecuteCount < toastabletool.ExecuteStart)
                    {   // Not executed enough yet
                        return false;
                    }
                    if (toastabletool.ExecuteEnd > 0 && tool.ExecuteCount > toastabletool.ExecuteEnd)
                    {   // Executed too many times, I give up
                        return false;
                    }
                    if ((tool.ExecuteCount - toastabletool.ExecuteStart) % toastabletool.ExecuteInterval != 0)
                    {   // Not the right execution interval
                        return false;
                    }
                    if (toastabletool.ExecutePercentChance < random.Next(1, 101))
                    {   // Random didn't want to toast
                        return false;
                    }
                    break;

                default:
                    return false;
            }
            try
            {
                ToastHelper.ToastIt(
                    plugin,
                    "supporting",
                    header: Settings.Instance.ToastHeader.Replace("{tool}", plugin.ToolName),
                    text: Settings.Instance.ToastText.Replace("{tool}", plugin.ToolName),
                    attribution: Settings.Instance.ToastAttrText.Replace("{tool}", plugin.ToolName),
                    logo: $"{Settings.URL}/Images/{tool.Acronym}150.png",
                    hero: $"{Settings.URL}/Images/SupportingHero.png",
                    buttons:
                    [
                        (Settings.Instance.ToastButtonCorporate.Replace("{tool}", plugin.ToolName), "corporate"),
                        (Settings.Instance.ToastButtonPersonal.Replace("{tool}", plugin.ToolName), "personal")
                    ]
                );
                tool.Support.ToastedDate = DateTime.Now;
                _ = installation.SaveAsync();
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void VerifySettings(string toolname, bool reload = false)
        {
            if (reload)
            {
                Settings.Reset();
                //RappPluginSettings.Instance.Save();    // this is only to get a correct format of the tool settings file
            }
            supportabletool = Settings.Instance[toolname];
            toastabletool = Settings.Instance.GetToastableTool(toolname);
        }

        private static void VerifyTool(string toolname, bool reload = false)
        {
            if (reload || installation == null || tool == null)
            {
                supporters = null;
                installation = Installation.Load();
                tool = installation[toolname];
                var version = Assembly.GetExecutingAssembly().GetName().Version;
                if (tool.version != version)
                {
                    tool.version = version;
                    tool.VersionRunDate = DateTime.Now;
                    _ = installation.SaveAsync();
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
                tool.Support.SubmittedDate.AddDays(Settings.Instance.ResetUnfinalizedSupportingAfterDays) < DateTime.Now &&
                supporters?.Any(s => s.Type == tool.Support.Type) == false)
            {
                tool.Support.Type = SupportType.None;
                tool.Support.SubmittedDate = DateTime.MinValue;
                _ = installation.SaveAsync();
            }
        }

        private static void OpenWebForm(string url, SupportType type)
        {
            tool.Support.Type = type;
            tool.Support.SubmittedDate = DateTime.Now;
            appinsights?.WriteEvent($"Supporting-{tool.Acronym}-{tool.Support.Type}");
            Process.Start(url);
        }

        #endregion Static Private Methods

        #region Private Constructors

        private Supporting(bool manual, SupportType? type)
        {
            InitializeComponent();
            lblHeader.Text = tool.Name;
            panInfo.Left = 32;
            panInfo.Top = 25;
            SetRandomPositions();
            SetStoredValues(manual);
            switch (type)
            {
                case SupportType.Company:
                    rbCompany.Checked = true;
                    break;

                case SupportType.Personal:
                    rbPersonal.Checked = true;
                    break;

                case SupportType.Contribute:
                    rbContribute.Checked = true;
                    break;
            }
        }

        #endregion Private Constructors

        #region Private Methods

        private void SetRandomPositions()
        {
            if (Settings.Instance.BMACLinkPositionRandom == true)
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
            if (Settings.Instance.CloseLinkPositionRandom == true)
            {
                var left = random.Next(0, 100);
                var top = random.Next(0, 100);
                if (left < 40)
                {
                    left = Settings.Instance.CloseLinkHorizFromOrigMin;
                }
                else if (left > 60)
                {
                    left = Settings.Instance.CloseLinkHorizFromOrigMax;
                }
                else
                {
                    left = (Settings.Instance.CloseLinkHorizFromOrigMin + Settings.Instance.CloseLinkHorizFromOrigMax) / 2;
                }

                if (top < 40)
                {
                    top = Settings.Instance.CloseLinkVertiFromOrigMin;
                }
                else if (top > 60)
                {
                    top = Settings.Instance.CloseLinkVertiFromOrigMax;
                }
                else
                {
                    top = (Settings.Instance.CloseLinkVertiFromOrigMin + Settings.Instance.CloseLinkVertiFromOrigMax) / 2;
                }

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
                    linkStatus.Text = Settings.Instance.StatusCompanyText.Replace("{tool}", tool.Name);
                    toolTip1.SetToolTip(linkStatus, Settings.Instance.StatusCompanyTip.Replace("{tool}", tool.Name));
                    break;

                case SupportType.Personal:
                    linkStatus.Text = Settings.Instance.StatusPersonalText.Replace("{tool}", tool.Name);
                    toolTip1.SetToolTip(linkStatus, Settings.Instance.StatusPersonalTip.Replace("{tool}", tool.Name));
                    break;

                case SupportType.Contribute:
                    linkStatus.Text = Settings.Instance.StatusContributeText.Replace("{tool}", tool.Name);
                    toolTip1.SetToolTip(linkStatus, Settings.Instance.StatusContributeTip.Replace("{tool}", tool.Name));
                    break;

                case SupportType.Already:
                    linkStatus.Text = Settings.Instance.StatusAlreadyText.Replace("{tool}", tool.Name);
                    toolTip1.SetToolTip(linkStatus, Settings.Instance.StatusAlreadyTip.Replace("{tool}", tool.Name));
                    break;

                case SupportType.Never:
                    linkStatus.Text = Settings.Instance.StatusNeverText.Replace("{tool}", tool.Name);
                    toolTip1.SetToolTip(linkStatus, Settings.Instance.StatusNeverTip.Replace("{tool}", tool.Name));
                    clickable = false;
                    break;

                default:
                    switch (tool.Support.Type)
                    {
                        case SupportType.Company:
                        case SupportType.Personal:
                        case SupportType.Contribute:
                        case SupportType.Already:
                            linkStatus.Text = Settings.Instance.StatusPendingText.Replace("{tool}", tool.Name);
                            toolTip1.SetToolTip(linkStatus, Settings.Instance.StatusPendingTip.Replace("{tool}", tool.Name));
                            clickable = false;
                            break;

                        case SupportType.Never:
                            linkStatus.Text = Settings.Instance.StatusNeverText.Replace("{tool}", tool.Name);
                            toolTip1.SetToolTip(linkStatus, Settings.Instance.StatusNeverTip.Replace("{tool}", tool.Name));
                            clickable = false;
                            break;

                        default:
                            linkStatus.Text = Settings.Instance.StatusDefaultText.Replace("{tool}", tool.Name);
                            toolTip1.SetToolTip(linkStatus, Settings.Instance.StatusDefaultTip.Replace("{tool}", tool.Name));
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
            panBgBlue.BackColor = Settings.Instance.clrBackground;
            panInfoBg.BackColor = Settings.Instance.clrBackground;
            helpText.BackColor = Settings.Instance.clrBackground;
            rbCompany.ForeColor = rbCompany.Checked ? Settings.Instance.clrTxtFgNormal : Settings.Instance.clrTxtFgDimmed;
            rbPersonal.ForeColor = rbPersonal.Checked ? Settings.Instance.clrTxtFgNormal : Settings.Instance.clrTxtFgDimmed;
            rbContribute.ForeColor = rbContribute.Checked ? Settings.Instance.clrTxtFgNormal : Settings.Instance.clrTxtFgDimmed;
            txtCompanyName.BackColor = Settings.Instance.clrFldBgNormal;
            txtCompanyEmail.BackColor = Settings.Instance.clrFldBgNormal;
            txtCompanyCountry.BackColor = Settings.Instance.clrFldBgNormal;
            txtPersonalFirst.BackColor = Settings.Instance.clrFldBgNormal;
            txtPersonalLast.BackColor = Settings.Instance.clrFldBgNormal;
            txtPersonalEmail.BackColor = Settings.Instance.clrFldBgNormal;
            txtPersonalCountry.BackColor = Settings.Instance.clrFldBgNormal;
            linkStatus.ForeColor = Settings.Instance.clrTxtFgDimmed;
            linkClose.LinkColor = Settings.Instance.clrTxtFgDimmed;
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
                //if (MessageBoxEx.Show(this, settings.ConfirmDirecting, "Supporting", MessageBoxButtons.OK, MessageBoxIcon.Asterisk) == DialogResult.OK)
                {
                    OpenWebForm(url, type);
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
                txtCompanyName.BackColor = string.IsNullOrEmpty(installation.CompanyName) ? Settings.Instance.clrFldBgInvalid : Settings.Instance.clrFldBgNormal;
            }
            if (sender == null || sender == txtCompanyEmail)
            {
                try
                {
                    installation.CompanyEmail = new MailAddress(txtCompanyEmail.Text).Address.Trim();
                }
                catch { }
                txtCompanyEmail.BackColor = string.IsNullOrEmpty(installation.CompanyEmail) ? Settings.Instance.clrFldBgInvalid : Settings.Instance.clrFldBgNormal;
            }
            if (sender == null || sender == txtCompanyCountry)
            {
                installation.CompanyCountry = txtCompanyCountry.Text.Trim().Length >= 2 ? txtCompanyCountry.Text.Trim() : "";
                txtCompanyCountry.BackColor = string.IsNullOrEmpty(installation.CompanyCountry) ? Settings.Instance.clrFldBgInvalid : Settings.Instance.clrFldBgNormal;
            }
            if (sender == null || sender == chkCompanySendInvoice)
            {
                installation.SendInvoice = chkCompanySendInvoice.Checked;
            }
            if (sender == null || sender == txtPersonalFirst)
            {
                installation.PersonalFirstName = txtPersonalFirst.Text.Trim().Length >= 1 ? txtPersonalFirst.Text.Trim() : "";
                txtPersonalFirst.BackColor = string.IsNullOrEmpty(installation.PersonalFirstName) ? Settings.Instance.clrFldBgInvalid : Settings.Instance.clrFldBgNormal;
            }
            if (sender == null || sender == txtPersonalLast)
            {
                installation.PersonalLastName = txtPersonalLast.Text.Trim().Length >= 2 ? txtPersonalLast.Text.Trim() : "";
                txtPersonalLast.BackColor = string.IsNullOrEmpty(installation.PersonalLastName) ? Settings.Instance.clrFldBgInvalid : Settings.Instance.clrFldBgNormal;
            }
            if (sender == null || sender == txtPersonalEmail)
            {
                try
                {
                    installation.PersonalEmail = new MailAddress(txtPersonalEmail.Text).Address.Trim();
                }
                catch { }
                txtPersonalEmail.BackColor = string.IsNullOrEmpty(installation.PersonalEmail) ? Settings.Instance.clrFldBgInvalid : Settings.Instance.clrFldBgNormal;
            }
            if (sender == null || sender == txtPersonalCountry)
            {
                installation.PersonalCountry = txtPersonalCountry.Text.Trim().Length >= 2 ? txtPersonalCountry.Text.Trim() : "";
                txtPersonalCountry.BackColor = string.IsNullOrEmpty(installation.PersonalCountry) ? Settings.Instance.clrFldBgInvalid : Settings.Instance.clrFldBgNormal;
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
            rbCompany.ForeColor = rbCompany.Checked ? Settings.Instance.clrTxtFgNormal : Settings.Instance.clrTxtFgDimmed;
            rbPersonal.ForeColor = rbPersonal.Checked ? Settings.Instance.clrTxtFgNormal : Settings.Instance.clrTxtFgDimmed;
            rbContribute.ForeColor = rbContribute.Checked ? Settings.Instance.clrTxtFgNormal : Settings.Instance.clrTxtFgDimmed;
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
                helpTitle.Text = Settings.Instance.HelpWhyTitle;
                helpText.Text = string.Empty;
                helpText.Text = Settings.Instance.HelpWhyText.Replace("\r\n", "\n").Replace("\n", "\r\n");
            }
            panInfo.Visible = visible;
        }

        private void btnInfo_Click(object sender, EventArgs e)
        {
            var visible = panInfo.Tag != sender || !panInfo.Visible;
            panInfo.Tag = sender;
            if (visible)
            {
                helpTitle.Text = Settings.Instance.HelpInfoTitle;
                helpText.Text = string.Empty;
                helpText.Text = Settings.Instance.HelpInfoText.Replace("\r\n", "\n").Replace("\n", "\r\n");
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

        private async void tsmiReset_Click(object sender, EventArgs e)
        {
            if (MessageBoxEx.Show(this, "Reset will remove all locally stored data regarding supporting.\nAnything submitted to Jonas will not be removed. If that is needed, please contact me directly.\n\nConfirm reset with Yes/No.", "Supporting", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
            {
                return;
            }
            var toolname = tool.Name;
            installation.Remove();
            VerifyTool(toolname, true);
            VerifySupporters(toolname, true);
            await installation.SaveAsync();
            SetStoredValues();
        }

        private void lbl_MouseEnter(object sender, EventArgs e)
        {
            if (sender is RadioButton rb)
            {
                rb.ForeColor = Settings.Instance.clrTxtFgNormal;
            }
            else if (sender is LinkLabel link)
            {
                link.LinkColor = Settings.Instance.clrTxtFgNormal;
            }
            else if (sender is Label lbl)
            {
                lbl.ForeColor = Settings.Instance.clrTxtFgNormal;
            }
        }

        private void lbl_MouseLeave(object sender, EventArgs e)
        {
            if (sender is RadioButton rb)
            {
                if (!rb.Checked)
                {
                    rb.ForeColor = Settings.Instance.clrTxtFgDimmed;
                }
            }
            else if (sender is LinkLabel link)
            {
                link.LinkColor = Settings.Instance.clrTxtFgDimmed;
            }
            else if (sender is Label lbl)
            {
                lbl.ForeColor = Settings.Instance.clrTxtFgDimmed;
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

    #region Supporters stored in Rappen.XTB.Supporters.xml

    public class Supporters : List<Supporter>
    {
        private const string SupportersFileName = "Rappen.XTB.Supporters.xml";

        public static Supporters DownloadMy(Guid InstallationId, string toolname, bool contributionCounts)
        {
            var result = XmlAtomicStore.DownloadXml<Supporters>(Settings.URL, SupportersFileName, Paths.SettingsPath);
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

    #endregion Supporters stored in Rappen.XTB.Supporters.xml

    public enum ShowItFrom
    {
        Open,
        Button,
        Execute,
        ToastCall,
        Other
    }
}