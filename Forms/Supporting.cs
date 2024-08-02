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
        private static RappenXTB tools;
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
                        Process.Start(url);
                    }
                    return;
                }
                if (reload || supporters == null)
                {
                    supporters = Supporters.DownloadMy(tools.InstallationId, toolname, supportabletool.ContributionCounts);
                }
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
                    else if (settings.ShowAutoPercentChance < 1 ||
                        settings.ShowAutoPercentChance <= random.Next(1, 100))
                    {
                        return;
                    }
                }
                if (manual && tool?.Support?.Type == SupportType.Never)
                {
                    tool.Support.Type = SupportType.None;
                }
                appinsights?.WriteEvent($"Supporting-{tool.Acronym}-Open-{(manual ? "Manual" : "Auto")}");
                new Supporting(manual).ShowDialog(plugin);
                if (!manual)
                {
                    tool.Support.AutoDisplayDate = DateTime.Now;
                    tool.Support.AutoDisplayCount++;
                }
                tools.Save();
            }
            catch (Exception ex)
            {
                plugin.LogError($"ToolSupporting error:\n{ex}");
            }
        }

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
            if (reload || tools == null || tool == null)
            {
                tools = RappenXTB.Load(settings);
                tool = tools[toolname];
                var version = Assembly.GetExecutingAssembly().GetName().Version;
                if (tool.version != version)
                {
                    tool.version = version;
                    tool.VersionRunDate = DateTime.Now;
                    tools.Save();
                }
            }
        }

        #endregion Static Public Methods

        #region Constructors

        private Supporting(bool manual)
        {
            InitializeComponent();
            lblHeader.Text = tool.Name;
            panInfo.Left = 32;
            panInfo.Top = 25;
            txtCompanyName.Text = tools.CompanyName;
            txtCompanyEmail.Text = tools.CompanyEmail;
            cmbCompanyUsers.SelectedIndex = tool.Support.UsersIndex;
            txtCompanyCountry.Text = tools.CompanyCountry;
            txtPersonalFirst.Text = tools.PersonalFirstName;
            txtPersonalLast.Text = tools.PersonalLastName;
            txtPersonalEmail.Text = tools.PersonalEmail;
            txtPersonalCountry.Text = tools.PersonalCountry;
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

        #endregion Constructors

        #region Private Methods

        private void SetAlreadyLink()
        {
            linkAlready.Tag = null;
            var supporter = supporters.OrderByDescending(s => s.Date).FirstOrDefault(s => s.Type != SupportType.None);
            switch (supporter?.Type)
            {
                case SupportType.Company:
                    linkAlready.Text = $"We're already\r\nsupporting\r\n{tool.Name}";
                    toolTip1.SetToolTip(linkAlready, $"We know that your company is supporting\r\n{tool.Name}\r\nThank You!");
                    break;

                case SupportType.Personal:
                    linkAlready.Text = $"I'm already\r\nsupporting\r\n{tool.Name}";
                    toolTip1.SetToolTip(linkAlready, $"We know that you are supporting\r\n{tool.Name}\r\nThank You!");
                    break;

                case SupportType.Contribute:
                    linkAlready.Text = $"I'm already\r\ncontributing to\r\n{tool.Name}";
                    toolTip1.SetToolTip(linkAlready, $"We know that you are contributing to\r\n{tool.Name}\r\nThank You!");
                    break;

                case SupportType.Already:
                    linkAlready.Text = $"I have already\r\nsupported\r\n{tool.Name}";
                    toolTip1.SetToolTip(linkAlready, $"We know that you have already supported\r\n{tool.Name}\r\nThank You!");
                    break;

                case SupportType.Never:
                    linkAlready.Text = $"I will never\r\nsupport\r\n{tool.Name}";
                    toolTip1.SetToolTip(linkAlready, $"For some strange reason,\r\nyou will never support\r\n{tool.Name}\r\nThink again?");
                    linkAlready.Tag = SupportType.Never;
                    break;

                case null:
                    linkAlready.Text = $"Register that\r\nI'm already\r\nsupporting";
                    toolTip1.SetToolTip(linkAlready, $"If you have already supported in any way to\r\n{tool.Name}\r\nClick here to let me know, and\r\nthis popup will not appear again!");
                    linkAlready.Tag = SupportType.Already;
                    break;
            }
            linkAlready.Visible = true;
        }

        private void ResetAllColors()
        {
            panBgBlue.BackColor = settings.clrBackground;
            panInfoBg.BackColor = settings.clrBackground;
            helpText.BackColor = settings.clrBackground;
            rbCompany.ForeColor = rbPersonal.Checked ? settings.clrTxtFgDimmed : settings.clrTxtFgNormal;
            rbPersonal.ForeColor = rbPersonal.Checked ? settings.clrTxtFgNormal : settings.clrTxtFgDimmed;
            rbPersonalSupporting.ForeColor = rbPersonalContributing.Checked ? settings.clrTxtFgDimmed : settings.clrTxtFgNormal;
            rbPersonalContributing.ForeColor = rbPersonalContributing.Checked ? settings.clrTxtFgNormal : settings.clrTxtFgDimmed;
            txtCompanyName.BackColor = settings.clrFldBgNormal;
            txtCompanyEmail.BackColor = settings.clrFldBgNormal;
            txtCompanyCountry.BackColor = settings.clrFldBgNormal;
            cmbCompanyUsers.BackColor = settings.clrFldBgNormal;
            txtPersonalFirst.BackColor = settings.clrFldBgNormal;
            txtPersonalLast.BackColor = settings.clrFldBgNormal;
            txtPersonalEmail.BackColor = settings.clrFldBgNormal;
            txtPersonalCountry.BackColor = settings.clrFldBgNormal;
            linkAlready.ForeColor = settings.clrTxtFgDimmed;
            linkClose.LinkColor = settings.clrTxtFgDimmed;
        }

        private void SettingAlready()
        {
            if (CallingWebForm(tool.GetUrlAlready()))
            {
                tool.Support.Type = SupportType.Already;
                DialogResult = DialogResult.Yes;
            }
        }

        private bool CallingWebForm(string url)
        {
            if (!string.IsNullOrEmpty(url))
            {
                if (MessageBoxEx.Show(this, settings.ConfirmDirecting, "Supporting", MessageBoxButtons.OK, MessageBoxIcon.Asterisk) == DialogResult.OK)
                {
                    tool.Support.Type = rbPersonal.Checked ? rbPersonalContributing.Checked ? SupportType.Contribute : SupportType.Personal : SupportType.Company;
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
                appinsights?.WriteEvent($"Supporting-{tool.Acronym}-Close", duration: sw.ElapsedMilliseconds);
            }
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            ctrl_Validating();
            var url = rbPersonal.Checked ? tool.GetUrlPersonal(rbPersonalContributing.Checked) : tool.GetUrlCorp();
            if (CallingWebForm(url))
            {
                DialogResult = DialogResult.Yes;
            }
        }

        private void ctrl_Validating(object sender = null, System.ComponentModel.CancelEventArgs e = null)
        {
            if (sender == null || sender == txtCompanyName)
            {
                tools.CompanyName = txtCompanyName.Text.Trim().Length >= 3 ? txtCompanyName.Text.Trim() : "";
                txtCompanyName.BackColor = string.IsNullOrEmpty(tools.CompanyName) ? settings.clrFldBgInvalid : settings.clrFldBgNormal;
            }
            if (sender == null || sender == txtCompanyEmail)
            {
                try
                {
                    tools.CompanyEmail = new MailAddress(txtCompanyEmail.Text).Address.Trim();
                }
                catch { }
                txtCompanyEmail.BackColor = string.IsNullOrEmpty(tools.CompanyEmail) ? settings.clrFldBgInvalid : settings.clrFldBgNormal;
            }
            if (sender == null || sender == txtCompanyCountry)
            {
                tools.CompanyCountry = txtCompanyCountry.Text.Trim().Length >= 2 ? txtCompanyCountry.Text.Trim() : "";
                txtCompanyCountry.BackColor = string.IsNullOrEmpty(tools.CompanyCountry) ? settings.clrFldBgInvalid : settings.clrFldBgNormal;
            }
            if (sender == null || sender == cmbCompanyUsers)
            {
                tool.Support.UsersIndex = cmbCompanyUsers.SelectedIndex;
                tool.Support.UsersCount = cmbCompanyUsers.Text;
                cmbCompanyUsers.BackColor = tool.Support.UsersIndex < 1 ? settings.clrFldBgInvalid : settings.clrFldBgNormal;
            }
            if (sender == null || sender == txtPersonalFirst)
            {
                tools.PersonalFirstName = txtPersonalFirst.Text.Trim().Length >= 1 ? txtPersonalFirst.Text.Trim() : "";
                txtPersonalFirst.BackColor = string.IsNullOrEmpty(tools.PersonalFirstName) ? settings.clrFldBgInvalid : settings.clrFldBgNormal;
            }
            if (sender == null || sender == txtPersonalLast)
            {
                tools.PersonalLastName = txtPersonalLast.Text.Trim().Length >= 2 ? txtPersonalLast.Text.Trim() : "";
                txtPersonalLast.BackColor = string.IsNullOrEmpty(tools.PersonalLastName) ? settings.clrFldBgInvalid : settings.clrFldBgNormal;
            }
            if (sender == null || sender == txtPersonalEmail)
            {
                try
                {
                    tools.PersonalEmail = new MailAddress(txtPersonalEmail.Text).Address.Trim();
                }
                catch { }
                txtPersonalEmail.BackColor = string.IsNullOrEmpty(tools.PersonalEmail) ? settings.clrFldBgInvalid : settings.clrFldBgNormal;
            }
            if (sender == null || sender == txtPersonalCountry)
            {
                tools.PersonalCountry = txtPersonalCountry.Text.Trim().Length >= 2 ? txtPersonalCountry.Text.Trim() : "";
                txtPersonalCountry.BackColor = string.IsNullOrEmpty(tools.PersonalCountry) ? settings.clrFldBgInvalid : settings.clrFldBgNormal;
            }
        }

        private void linkClose_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            DialogResult = DialogResult.Retry;
        }

        private void linkAlready_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var type = linkAlready.Tag as SupportType?;
            switch (type)
            {
                case SupportType.Already:
                    SettingAlready();
                    break;

                case SupportType.Never:
                    MessageBoxEx.Show("You can change your mind right now or later.", "Supporting");
                    break;

                default:
                    MessageBoxEx.Show("Thanks! ❤️", "Supporting");
                    break;
            }
        }

        private void rbType_CheckedChanged(object sender, EventArgs e)
        {
            SuspendLayout();
            rbCompany.ForeColor = rbPersonal.Checked ? settings.clrTxtFgDimmed : settings.clrTxtFgNormal;
            rbPersonal.ForeColor = rbPersonal.Checked ? settings.clrTxtFgNormal : settings.clrTxtFgDimmed;
            panPersonal.Left = panCorp.Left;
            panPersonal.Top = panCorp.Top;
            panPersonal.Visible = rbPersonal.Checked;
            panCorp.Visible = !panPersonal.Visible;
            btnSubmit.ImageIndex = rbPersonal.Checked ? rbPersonalContributing.Checked ? 2 : 1 : 0;
            ResumeLayout();
        }

        private void rbPersonalMonetary_CheckedChanged(object sender, EventArgs e)
        {
            rbPersonalSupporting.ForeColor = rbPersonalContributing.Checked ? settings.clrTxtFgDimmed : settings.clrTxtFgNormal;
            rbPersonalContributing.ForeColor = rbPersonalContributing.Checked ? settings.clrTxtFgNormal : settings.clrTxtFgDimmed;
            btnSubmit.ImageIndex = rbPersonalContributing.Checked ? 2 : 1;
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
            SettingAlready();
        }

        private void tsmiNever_Click(object sender, EventArgs e)
        {
            tool.Support.Type = SupportType.Never;
            DialogResult = DialogResult.Yes;
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

        #endregion Private Event Methods
    }

    public class ToolSettings
    {
        private const string FileName = "Rappen.XTB.Settings.xml";
        private static readonly Uri ToolSettingsURLPath = new Uri("https://raw.githubusercontent.com/rappen/Rappen.XTB.Supporting/main/Config/");

        public int SettingsVersion = 1;
        public List<SupportableTool> SupportableTools = new List<SupportableTool>();
        public int ShowMinutesAfterToolInstall = int.MaxValue;    // 60
        public int ShowMinutesAfterToolNewVersion = int.MaxValue; // 120
        public int ShowMinutesAfterSupportingShown = int.MaxValue; // 2880m / 48h / 2d
        public int ShowMinutesAfterSubmitting = int.MaxValue; // 2880m / 48h / 2d
        public int ShowAutoPercentChance = 0;   // 25 (0-100)
        public int ShowAutoRepeatTimes = 0; // 10

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
            "&{formid}_19={size}" +
            "&{formid}_24={amount}" +
            "&{formid}_27={company}" +
            "&{formid}_31={tool}" +
            "&{formid}_32={version}" +
            "&{formid}_33={instid}";

        public string FormUrlSupporting =
            "https://jonasr.app/supporting/personal-prefilled/" +
            "?{formid}_1_first={firstname}" +
            "&{formid}_1_last={lastname}" +
            "&{formid}_3={country}" +
            "&{formid}_4={email}" +
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

        public string ColorBg = "FF0042AD";
        public string ColorFieldBgNormal = "FF0063FF";
        public string ColorFieldBgInvalid = "FFF06565";
        public string ColorTextFgNormal = "FFFFFF00";
        public string ColorTextFgDimmed = "FFD2B48C";

        public Color clrBackground => Color.FromArgb(int.Parse(ColorBg, System.Globalization.NumberStyles.HexNumber));
        public Color clrTxtFgNormal => Color.FromArgb(int.Parse(ColorTextFgNormal, System.Globalization.NumberStyles.HexNumber));
        public Color clrTxtFgDimmed => Color.FromArgb(int.Parse(ColorTextFgDimmed, System.Globalization.NumberStyles.HexNumber));
        public Color clrFldBgNormal => Color.FromArgb(int.Parse(ColorFieldBgNormal, System.Globalization.NumberStyles.HexNumber));
        public Color clrFldBgInvalid => Color.FromArgb(int.Parse(ColorFieldBgInvalid, System.Globalization.NumberStyles.HexNumber));

        public string ConfirmDirecting = @"You will now be redirected to the website form
to finish Your flavor of support.
After the form is submitted, Jonas will handle it soon.

NOTE: It has to be submitted during the next step!";

        public string HelpWhyTitle = "Community Tools are Conscienceware.";

        public string HelpWhyText = @"Some of us in the Power Platform Community are creating tools.
Some contribute to the community with new ideas, find problems, write documentation, and even solve our bugs.
Thousands and thousands in this community are mostly 'consumers'—only using open-source tools.
To me, it's very similar to watching TV. Do you pay for channels, Netflix, Amazon Prime, Spotify, etc.?
To be part of the community, but without the examples above, you can simply pay instead.

Especially when you work in a big corporation, exploiting free tools - only to increase your income - you have a responsibility to participate actively in the community - or pay.
It's good to be able to sleep with a good conscience. Right?

There should be a license called ""Conscienceware"".
But technically, it is simply free to use them.

If you say you are not part of the community, that is incorrect—just using these tools makes you a part of it.

You and your company can now more formally support tools rather than just donating via PayPal or 'Buy Me a Coffee.'

Supporting is not just giving money; it means that you or your company know you have gained in time and improved your quality by using these tools. If you get something and want to give back—support the development and maintenance of the tools.

To read more about my thoughts, click here: https://jonasr.app/helping/

- Jonas Rapp";

        public string HelpInfoTitle = "Technical Information";

        public string HelpInfoText = @"Your entered name, company, country, email, and amount will not be stored in any system. The information will be saved in my personal Excel file. I do this to ensure you can get an invoice, and if so, we need to communicate if necessary.
The email you share with me, only to me, will never be sold to any company.

You will receive an official receipt immediately and, if needed, an invoice. Supporting can be done with a credit card. Other options will be available depending on your location. Stripe handles the payment.

When you click the big button here, the information you entered here will be included in the form on my website, jonasr.app, and a few hidden info: tool name, version, and your XrmToolBox 'InstallationId' (a random Guid generated the first time you use the toolbox). If you are curious, you can find your ID here: https://jonasr.app/xtb-finding-installationid.

Since I would like to be very clear and transparent - we store your XrmToolBox InstallationId on a server to be able to know that you are supporting it in some way. There is nothing about the amount or contribution; I am not interested in hacking this info.

The button in the top-right corner opens this info. You can also right-click on it and find more options, especially:
* I have already supported this tool — use this to tell me that you already support this tool in some way so that this prompt will not ask you again.
* I will never support this tool — use it if you think it is a bad idea, and you probably won't use it again; it won't ask you again.

For questions, contact me at https://jonasr.app/contact.";

        private ToolSettings()
        { }

        public static ToolSettings Get() => new Uri(ToolSettingsURLPath, FileName).DownloadXml<ToolSettings>() ?? new ToolSettings();

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
    }

    public class SupportableTool
    {
        public string Name;
        public bool Enabled = false;
        public bool ShowAutomatically = false;
        public bool ContributionCounts = true;
    }

    public class Supporters : List<Supporter>
    {
        private const string FileName = "Rappen.XTB.Supporters.xml";
        private static readonly Uri SupportersURLPath = new Uri("https://raw.githubusercontent.com/rappen/Rappen.XTB.Supporting/main/Config/");

        public static Supporters DownloadMy(Guid InstallationId, string toolname, bool contributionCounts)
        {
            var result = new Uri(SupportersURLPath, FileName).DownloadXml<Supporters>() ?? new Supporters();
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

    public class RappenXTB
    {
        private int settingversion = -1;
        internal ToolSettings toolsettings;

        public int SettingsVersion
        {
            get => settingversion;
            set
            {
                if (settingversion != -1 && value != settingversion && Tools?.Count() > 0)
                {
                    Tools.ForEach(s => s.Support.AutoDisplayCount = 0);
                }
                settingversion = value;
            }
        }

        public Guid InstallationId = Guid.Empty;
        public string CompanyName;
        public string CompanyEmail;
        public string CompanyCountry;
        public string PersonalFirstName;
        public string PersonalLastName;
        public string PersonalEmail;
        public string PersonalCountry;
        public List<Tool> Tools = new List<Tool>();

        public static RappenXTB Load(ToolSettings settings)
        {
            string path = Path.Combine(Paths.SettingsPath, "Rappen.XTB.Tools.xml");
            var result = new RappenXTB();
            if (File.Exists(path))
            {
                try
                {
                    XmlDocument xmlDocument = new XmlDocument();
                    xmlDocument.Load(path);
                    result = (RappenXTB)XmlSerializerHelper.Deserialize(xmlDocument.OuterXml, typeof(RappenXTB));
                }
                catch { }
            }
            result.toolsettings = settings;
            if (result.InstallationId.Equals(Guid.Empty))
            {
                result.InstallationId = InstallationInfo.Instance.InstallationId;
            }
            if (settings.SettingsVersion != result.settingversion)
            {
                result.SettingsVersion = settings.SettingsVersion;
            }
            result.Tools.ForEach(t => t.RappenXTB = result);
            return result;
        }

        public void Save()
        {
            if (!Directory.Exists(Paths.SettingsPath))
            {
                Directory.CreateDirectory(Paths.SettingsPath);
            }
            string path = Path.Combine(Paths.SettingsPath, "Rappen.XTB.Tools.xml");
            XmlSerializerHelper.SerializeToFile(this, path);
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

        internal RappenXTB RappenXTB;

        internal Version version
        {
            get => _version;
            set
            {
                if (value != _version)
                {
                    Support.AutoDisplayCount = 0;
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
                    case "FetchXML Builderx":
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

        internal Tool(RappenXTB rappen, string name)
        {
            RappenXTB = rappen;
            Name = name;
        }

        public string GetUrlCorp()
        {
            if (string.IsNullOrEmpty(RappenXTB.CompanyName) ||
                string.IsNullOrEmpty(RappenXTB.CompanyEmail) ||
                string.IsNullOrEmpty(RappenXTB.CompanyCountry) ||
                Support.UsersIndex < 1)
            {
                return null;
            }
            return GenerateUrl(RappenXTB.toolsettings.FormUrlCorporate, RappenXTB.toolsettings.FormIdCorporate);
        }

        public string GetUrlPersonal(bool contribute)
        {
            if (string.IsNullOrEmpty(RappenXTB.PersonalFirstName) ||
                string.IsNullOrEmpty(RappenXTB.PersonalLastName) ||
                string.IsNullOrEmpty(RappenXTB.PersonalEmail) ||
                string.IsNullOrEmpty(RappenXTB.PersonalCountry))
            {
                return null;
            }
            return GenerateUrl(contribute ? RappenXTB.toolsettings.FormUrlContribute : RappenXTB.toolsettings.FormUrlSupporting, contribute ? RappenXTB.toolsettings.FormIdContribute : RappenXTB.toolsettings.FormIdPersonal);
        }

        public string GetUrlAlready()
        {
            return GenerateUrl(RappenXTB.toolsettings.FormUrlAlready, RappenXTB.toolsettings.FormIdAlready);
        }

        public string GetUrlGeneral()
        {
            return GenerateUrl(RappenXTB.toolsettings.FormUrlGeneral, RappenXTB.toolsettings.FormIdCorporate);
        }

        private string GenerateUrl(string template, string form)
        {
            return template
                .Replace("{formid}", form)
                .Replace("{company}", RappenXTB.CompanyName)
                .Replace("{invoiceemail}", RappenXTB.CompanyEmail)
                .Replace("{companycountry}", RappenXTB.CompanyCountry)
                .Replace("{amount}", Support.Amount)
                .Replace("{size}", Support.UsersCount)
                .Replace("{firstname}", RappenXTB.PersonalFirstName)
                .Replace("{lastname}", RappenXTB.PersonalLastName)
                .Replace("{email}", RappenXTB.PersonalEmail)
                .Replace("{country}", RappenXTB.PersonalCountry)
                .Replace("{tool}", Name)
                .Replace("{version}", version.ToString())
                .Replace("{instid}", RappenXTB.InstallationId.ToString());
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