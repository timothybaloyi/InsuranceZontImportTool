using Microsoft.Crm.Sdk.Messages;
using Microsoft.Win32;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace InsuranceZontImportTool
{
    /// <summary>
    /// Interaction logic for ImportWindow.xaml
    /// </summary>
    public partial class ImportWindow : Window
    {
        public ImportWindow()
        {
            InitializeComponent();
        }

        List<CorpInfo> _corpInfos = new List<CorpInfo>();
        List<Contact> _contacts = new List<Contact>();
        List<FundMember> _fundMembers = new List<FundMember>();
        List<LeadRegister> _leadRegister = new List<LeadRegister>();


        List<Guid> _newLeadIds = new List<Guid>();
        List<Guid> _newCaseIds = new List<Guid>();

        public void BindData(List<CorpInfo> data)
        {
            _corpInfos = data;

            foreach (var corp in data)
            {
                ConvertCorpDataToObjects(corp);
            }

            int advisorCount = data.GroupBy(x => x.Advisor).Count();
            lblAdvisors.Content = advisorCount;

            int fundCount = data.GroupBy(x => x.Fund).Count();
            lblFundsNum.Content = fundCount;

            int memberCount = (from r in _fundMembers
                               group r by new
                               {
                                   r.MemberID,
                                   r.MemberNo
                               } into g
                               select g
                               ).Count();

            lblMembersNum.Content = memberCount;

        }

        private void ConvertCorpDataToObjects(CorpInfo corp)
        {
            foreach (DataRow row in corp.Data.Rows)
            {
                Guid tempID = Guid.NewGuid();

                //Title,Surname,First Name,Date of Birth,Gender,Marital Status,Tel No (Home),Tel No (Work),Cellular Tel No,E-mail address,
                _contacts.Add(new Contact()
                {
                    Advisor = corp.Advisor,
                    Cellular = row["Cellular Tel No"].ToString(),
                    DateOfBirth = row["Date of Birth"].ToString(),
                    Email = row["E-mail address"].ToString(),
                    FirstName = row["First Name"].ToString(),
                    Fund = corp.Fund,
                    Gender = row["Gender"].ToString(),
                    IDNo = row["Member ID No"].ToString(),
                    LastName = row["Surname"].ToString(),
                    MaritalStatus = row["Marital Status"].ToString(),
                    MemberNo = row["Member No"].ToString(),
                    TelNoHome = row["Tel No (Home)"].ToString(),
                    TelNoWork = row["Tel No (Work)"].ToString(),
                    TempLinkID = tempID,
                    Title = row["Title"].ToString()

                });

                //Date member joined employer,Joined take over scheme date,Total Remuneration,Risk Salary,RA Monthly Premium,Pay frequency,Number of months/weeks to annualise,Member status,Member ID No,Member ID Type,Payroll No,Member No,Member Paypoint Name,
                _fundMembers.Add(new FundMember()
                {
                    Advisor = corp.Advisor,
                    DateMemberJoinedEmployer = row["Date member joined employer"].ToString(),
                    Fund = corp.Fund,
                    TotalRemuneration = row["Total Remuneration"].ToString(),
                    TempLinkID = tempID,
                    JoinedTakeOverSchemDate = row["Joined take over scheme date"].ToString(),
                    MemberID = row["Member ID No"].ToString(),
                    MemberIDType = row["Member ID Type"].ToString(),
                    MemberNo = row["Member No"].ToString(),
                    MemberPaypointName = row["Member Paypoint Name"].ToString(),
                    MemberStatus = row["Member status"].ToString(),
                    NumberOfMonthsWeeksToAnnualise = row["Number of months/weeks to annualise"].ToString(),
                    PayFrequency = row["Pay frequency"].ToString(),
                    PayrollNo = row["Payroll No"].ToString(),
                    RAMonthlyPremium = row["RA Monthly Premium"].ToString(),
                    RiskSalary = row["Risk Salary"].ToString()
                });
            }
        }


        private void btnLeadsRegister_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDlg = new OpenFileDialog();
            if (fileDlg.ShowDialog().Value)
            {
                string registerPath = fileDlg.FileName;

                LoadDataTable(registerPath, "Lead Register");

                if (_leadRegister.Count > 0)
                {
                    FilterOutNewLeads();

                    LeadRegisterToCRM();
                }
            }
        }

        private void LeadRegisterToCRM()
        {
            List<string> leadStatuses = new List<string>()
            {
                "",
                "IN PROGRESS".Trim(),
                "NEW ALLOCATION ".Trim(),
                "in progress ".Trim(),
                "In Contact ".Trim(),
                "Haven't contacted yet ".Trim()
            };

            foreach (var item in _leadRegister)
            {
                
                Guid tempId = Guid.NewGuid();

                _contacts.Add(new Contact()
                {
                    Advisor = item.ISSUED_TO_CONSULTANT,
                    Cellular = item.CONATACT_DETAILS,
                    Email = item.EMAIL,
                    FirstName = item.FIRST_NAME,
                    Fund = item.FUND_NAME,
                    IDNo = item.ID_NUMBER,
                    Gender = item.GENDER,
                    LastName = item.SURNAME,
                    MemberNo = item.MEMBER_NUMBER,
                    TempLinkID = tempId
                });

                _fundMembers.Add(new FundMember()
                {
                    TempLinkID = tempId,
                    Fund = item.FUND_NAME,
                    Advisor = item.ISSUED_TO_CONSULTANT,
                    MemberNo = item.MEMBER_NUMBER,
                    MemberID = item.ID_NUMBER,
                    RiskSalary = item.RISK_SALARY
                });

                if (leadStatuses.Contains(item.Status.Trim()))
                {
                    _newLeadIds.Add(tempId);
                }
                else
                {
                    _newCaseIds.Add(tempId);
                }

            }
        }

        private void FilterOutNewLeads()
        {
            //Records that in fundmember but not in the Lead Register
            int newLeads = 0;

            foreach (var item in _contacts)
            {
                bool leadFound = _leadRegister.Where(x => x.FIRST_NAME.Equals(item.FirstName, StringComparison.InvariantCultureIgnoreCase) && x.SURNAME.Equals(item.LastName, StringComparison.InvariantCultureIgnoreCase) && x.MEMBER_NUMBER.Equals(item.MemberNo, StringComparison.InvariantCultureIgnoreCase)).Any();
                if (!leadFound)
                {
                    _newLeadIds.Add(item.TempLinkID);
                    newLeads++;
                }
            }

            lblLeadsNum.Content = newLeads;
        }

        private void LoadDataTable(string path, string sheetName)
        {
            OleDbConnection cnn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0 Xml");

            OleDbCommand oconn = new OleDbCommand($"select * from [{sheetName}$]", cnn);
            cnn.Open();

            OleDbDataReader reader = oconn.ExecuteReader();
            bool isHeader = true;
            DataTable dt = new DataTable();
            int fields = 0;
            int usedFields = 0;
            while (reader.Read())
            {
                if (isHeader)
                {
                    isHeader = false;
                    fields = reader.FieldCount;
                    for (int i = 0; i < fields; i++)
                    {
                        string fieldName = reader[i].ToString();
                        fieldName = fieldName.Trim();

                        if (!dt.Columns.Contains(fieldName))
                        {
                            usedFields++;
                            dt.Columns.Add(fieldName);
                        }
                        else
                        {
                            usedFields++;
                            dt.Columns.Add(fieldName + i);
                        }
                    }
                }
                else
                {
                    DataRow newRow = dt.NewRow();
                    for (int i = 0; i < usedFields; i++)
                    {
                        newRow[i] = reader[i];
                    }
                    dt.Rows.Add(newRow);
                }
            }

            if (dt != null)
            {

                foreach (DataRow row in dt.Rows)
                {
                    /*
                     'ISSUED TO CONSULTANT'
                    'LEAD SOURCE'
                    'LEAD ISSUE'
                    'FUND NAME'
                    'BPC INFO'
                    'FUND VALUE'
                    'SURNAME'
                    'FIRST NAME'
                    ''
                    'GENDER'
                    'MEMBER NUMBER'
                    'ID NUMBER'
                    'EMAIL'
                    'LOCATION'
                    ''
                    'Last date seen/made contact with'
                    'last date seen / made contact'
                    'Status'
                    ''
                    ''
                    ''
                    'REALLOCATION FROM'
                    'REALLOCATED FROM'
                    'REALLOCATION DATE'
                    'REALLOCATIO FROM'
                    'NOTES' 
                     */
                    _leadRegister.Add(new LeadRegister()
                    {
                        BPC_INFO = row["BPC INFO"].ToString(),
                        Additional_Information = row["Additional Information"].ToString(),
                        CONATACT_DETAILS = row["CONATACT DETAILS"].ToString(),
                        EMAIL = row["EMAIL"].ToString(),
                        FIRST_NAME = row["FIRST NAME"].ToString(),
                        FNA_DATE = row["FNA DATE"].ToString(),
                        FNA_Presented = row["FNA  Presented"].ToString(),
                        FUND_NAME = row["FUND NAME"].ToString(),
                        GENDER = row["GENDER"].ToString(),
                        ID_NUMBER = row["ID NUMBER"].ToString(),
                        ISSUED_TO_CONSULTANT = row["ISSUED TO CONSULTANT"].ToString(),
                        Last_date_seenmade_contact_with = row["Last date seen/made contact with"].ToString(),
                        LEAD_ISSUE = row["LEAD ISSUE"].ToString(),
                        LEAD_SOURCE = row["LEAD SOURCE"].ToString(),
                        LOCATION = row["LOCATION"].ToString(),
                        MEMBER_NUMBER = row["MEMBER NUMBER"].ToString(),
                        NOTES = row["NOTES"].ToString(),
                        REALLOCATION_DATE = row["REALLOCATION DATE"].ToString(),
                        REALLOCATION_FROM = row["REALLOCATION FROM"].ToString(),
                        SECOND_NAME = row["2ND NAME"].ToString(),
                        Status = row["Status"].ToString(),
                        SURNAME = row["SURNAME"].ToString(),
                    });
                }
            }

        }

        private void ImportContacts()
        {
            CrmServiceClient conn = new Microsoft.Xrm.Tooling.Connector.CrmServiceClient(ConfigurationManager.ConnectionStrings["CRM Online"].ConnectionString);

            // Cast the proxy client to the IOrganizationService interface.
            IOrganizationService _orgService = (IOrganizationService)conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;

            foreach (var item in _contacts)
            {
                Entity contact = new Entity("contact");
                contact["firstname"] = item.FirstName;
                contact["lastname"] = item.LastName;
                contact["governmentid"] = item.IDNo;
                
            }
            


        }

        private void ImportFunds()
        {
           
        }

        private void ImportFundMembers()
        {

        }

        private void ImportLeads()
        {

        }

        private void ImportCases()
        {

        }

        private void btnImportINCRM_Click(object sender, RoutedEventArgs e)
        {
            ImportContacts();
        }
    }

    //Title,Surname,Maiden Name,First Name,Date of Birth,Gender,Marital Status,Tel No (Home),Tel No (Work),Cellular Tel No,E-mail address,

    public class Contact
    {
        public Guid TempLinkID { get; set; }

        public string Advisor { get; set; }

        public string Fund { get; set; }

        public string Title { get; set; }

        public string FirstName { get; set; }

        public string LastName { get; set; }

        public string IDNo { get; set; }

        public string MemberNo { get; set; }

        public string DateOfBirth { get; set; }

        public string Gender { get; set; }

        public string MaritalStatus { get; set; }

        public string TelNoHome { get; set; }

        public string TelNoWork { get; set; }

        public string Cellular { get; set; }

        public string Email { get; set; }

    }

    //Date member joined employer,Joined take over scheme date,Total Remuneration,Risk Salary,RA Monthly Premium,Pay frequency,Number of months/weeks to annualise,Member status,Member ID No,Member ID Type,Payroll No,Member No,Member Paypoint Name,
    public class FundMember
    {
        public Guid TempLinkID { get; set; }

        public string Advisor { get; set; }

        public string Fund { get; set; }

        public string DateMemberJoinedEmployer { get; set; }

        public string JoinedTakeOverSchemDate { get; set; }

        public string TotalRemuneration { get; set; }

        public string RiskSalary { get; set; }

        public string RAMonthlyPremium { get; set; }

        public string PayFrequency { get; set; }

        public string NumberOfMonthsWeeksToAnnualise { get; set; }

        public string MemberStatus { get; set; }

        public string MemberID { get; set; }

        public string MemberNo { get; set; }

        public string PayrollNo { get; set; }

        public string MemberPaypointName { get; set; }

        public string MemberIDType { get; set; }

    }

    public class LeadRegister
    {

        public string ISSUED_TO_CONSULTANT { get; set; }
        public string LEAD_SOURCE { get; set; }
        public string LEAD_ISSUE { get; set; }
        public string ISSUE_DATE { get; set; }
        public string FUND_NAME { get; set; }
        public string SURNAME { get; set; }
        public string FIRST_NAME { get; set; }
        public string SECOND_NAME { get; set; }
        public string GENDER { get; set; }
        public string MEMBER_NUMBER { get; set; }
        public string BPC_INFO { get; set; }
        public string ID_NUMBER { get; set; }
        public string Normal_Retirement_Date { get; set; }
        public string RISK_SALARY { get; set; }
        public string EMAIL { get; set; }
        public string LOCATION { get; set; }
        public string CONATACT_DETAILS { get; set; }
        public string STATUS_DATE { get; set; }
        public string Last_date_seenmade_contact_with { get; set; }
        public string Last_date_seenmade_contact_with1 { get; set; }
        public string Last_date_seenmade_contact_with2 { get; set; }
        public string Status { get; set; }
        public string FNA_Presented { get; set; }
        public string FNA_DATE { get; set; }
        public string Additional_Information { get; set; }
        public string REALLOCATION_FROM { get; set; }
        public string REALLOCATION_DATE { get; set; }
        public string REALLOCATED_FROM1 { get; set; }
        public string REALLOCATION_DATE1 { get; set; }
        public string REALLOCATIO_FROM2 { get; set; }
        public string REALLOCATION_DATE2 { get; set; }
        public string REALLOCATIO_FROM3 { get; set; }
        public string REALLOCATION_DATE3 { get; set; }
        public string NOTES { get; set; }

    }
}
