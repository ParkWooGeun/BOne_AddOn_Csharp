using Microsoft.VisualBasic;
using System;
using System.Windows.Forms;
using SAPbouiCOM;
using Scripting;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// MainClass : �������� �ʱ�ȭ, ���ø����̼� ����, DI API ����, ȸ�� DB ����, ODBC ����� ���� �ʱ�ȭ, MainMenu�� XML �ε�, ��ȿ�� �˻�, AddOn �� ����, System �� ����, �̺�Ʈ ����, �̺�Ʈ ���� ����
    /// ZZMDC Ŭ������ ��Ī
    /// </summary>
    internal class PSH_MainClass
    {
        /// <summary>
        /// ������
        /// </summary>
        public PSH_MainClass() : base()
        {
            this.Initialize_Calss(); //Ŭ���� �ʱ�ȭ

            //�̺�Ʈ ����
            PSH_Globals.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            PSH_Globals.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            PSH_Globals.SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
            PSH_Globals.SBO_Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);
            PSH_Globals.SBO_Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);
        }

        /// <summary>
        /// Ŭ���� �ʱ�ȭ
        /// </summary>
        private void Initialize_Calss()
        {
            try
            {
                this.Initialize_GlobalVariable();
                this.Connect_Application();

                // Set The Connection Context
                if (!(Connect_DIAPI() == 0))
                {
                    PSH_Globals.SBO_Application.MessageBox("DI API �������", 1, "Ok", "", "");
                    System.Environment.Exit(0);
                }

                // Connect To The Company Data Base
                if (!(Connect_CompanyDB() == 0))
                {
                    PSH_Globals.SBO_Application.MessageBox("ȸ�� DB �������", 1, "Ok", "", "");
                    System.Environment.Exit(0);
                }

                PSH_SetFilter.Execute(); //Event Filter Excute
                //PSH_EventHelpClass eventHelpClass = new PSH_EventHelpClass();
                //PSH_BaseClass baseClass = new PSH_BaseClass();
                //eventHelpClass.Set_EventFilter(baseClass);

                this.XmlCreateYN();
                this.Load_MenuXml();
                //DoSomething();

                Initialize_ODBC_Variable();

                PSH_Globals.SBO_Application.StatusBar.SetText("PSH_BOne_AddOn �ʱ�ȭ �Ϸ�", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Initialize_Calss_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// �������� �ʱ�ȭ
        /// </summary>
		private void Initialize_GlobalVariable()
        {
            PSH_Globals.FormCurrentCount = 0;
            PSH_Globals.FormTotalCount = 0;
            PSH_Globals.ClassList = new Collection();
            PSH_Globals.FormTypeListCount = 0;
            PSH_Globals.FormTypeList = new Collection();
            PSH_Globals.oCompany = new SAPbobsCOM.Company();
            PSH_Globals.Screen = "Screen";
            PSH_Globals.Report = "Report";
        }

        /// <summary>
        /// ���ø����̼� ����
        /// </summary>
        private void Connect_Application()
        {
            try
            {
                SAPbouiCOM.SboGuiApi SboGuiApi = new SAPbouiCOM.SboGuiApi();

                string ConnectionString = string.Empty;

                ConnectionString = Interaction.Command();

                if (string.IsNullOrEmpty(Strings.Trim(ConnectionString)))
                {
                    ConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
                }

                SboGuiApi.Connect(ConnectionString);
                PSH_Globals.SBO_Application = SboGuiApi.GetApplication(-1);
                PSH_Globals.SBO_Application.StatusBar.SetText("PSH_BOne_AddOn ���� ��...", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("PSH_BOne_AddOn ���� ���� : " + ex.Message, "SAP Business One", MessageBoxButtons.YesNo);
            }
        }

        /// <summary>
        /// DI API ����
        /// </summary>
        /// <returns>0 : ����</returns>
        private int Connect_DIAPI()
        {
            int setConnectionContextReturn = 0;

            string sCookie = string.Empty;
            string sConnectionContext = string.Empty;

            // acquire the connection context cookie from the DI API
            sCookie = PSH_Globals.oCompany.GetContextCookie();

            // retrieve the connection context string from the UI API using the acquired cookie
            sConnectionContext = PSH_Globals.SBO_Application.Company.GetConnectionContext(sCookie);

            // before setting the SBO Login Context make sure the company is not connected
            if (PSH_Globals.oCompany.Connected == true)
            {
                PSH_Globals.oCompany.Disconnect();
            }

            // Set the connection context information to the DI API
            setConnectionContextReturn = PSH_Globals.oCompany.SetSboLoginContext(sConnectionContext);

            return setConnectionContextReturn;
        }

        /// <summary>
        /// ȸ�� DB ����
        /// </summary>
        /// <returns>0 : ����</returns>
        private int Connect_CompanyDB()
        {
            int connectToCompanyReturn = 0;

            // Establish the connection to the company database.
            connectToCompanyReturn = PSH_Globals.oCompany.Connect();

            return connectToCompanyReturn; //36,000ms ~ 40,000ms �ҿ�
        }

        /// <summary>
        /// ODBC ����� ���� �ʱ�ȭ
        /// </summary>
        public void Initialize_ODBC_Variable()
        {
            string sQry = string.Empty;
            string ServerName = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = null;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            oRecordSet = (SAPbobsCOM.Recordset)PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            ServerName = PSH_Globals.SBO_Application.Company.ServerName;

            sQry = "        SELECT      PARAM01 AS PARAM01,";
            sQry = sQry + "             PARAM02 AS PARAM02,";
            sQry = sQry + "             PARAM03 AS PARAM03,";
            sQry = sQry + "             PARAM04 AS PARAM04,";
            sQry = sQry + "             PARAM05 AS PARAM05,";
            sQry = sQry + "             PARAM06 AS PARAM06,";
            sQry = sQry + "             PARAM07 AS PARAM07,";
            sQry = sQry + "             PARAM08 AS PARAM08";
            sQry = sQry + " FROM        PROFILE ";
            sQry = sQry + " WHERE       TYPE = 'SERVERINFO'";

            oRecordSet.DoQuery(sQry);

            if (oRecordSet.RecordCount > 0)
            {
                //ODBC
                //PSH_Globals.SP_ODBC_YN = Trim(oRecordset.Fields("Value01").Value)
                if (codeHelpClass.Right(ServerName, 3) == "223"){
                    PSH_Globals.SP_ODBC_Name = "MDCERP";
                }
                else
                {
                    PSH_Globals.SP_ODBC_Name = "PSHERP_TEST"; // 191.1.1.223���� ���ӽ� ����  ODBC�� ����
                }
                PSH_Globals.SP_ODBC_IP = ServerName;
                //������ �����ּҸ� �ٷ� �����ü� �ְ� ���� PARAM01���� ������ ���� PSH_Globals.SBO_Application.Company.ServerName
                //PSH_Globals.SP_ODBC_IP = oRecordSet.Fields.Item("PARAM01").Value.ToString().Replace("\\", "").Trim();
                PSH_Globals.SP_ODBC_DBName = PSH_Globals.oCompany.CompanyDB;
                PSH_Globals.SP_ODBC_ID = oRecordSet.Fields.Item("PARAM07").Value.ToString().Trim();
                PSH_Globals.SP_ODBC_PW = oRecordSet.Fields.Item("PARAM08").Value.ToString().Trim();
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //�޸� ����
        }

        /// <summary>
        /// ���� �޴��� XML �ε�
        /// </summary>
        private void XmlCreateYN()
        {
            string Query01 = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            FileSystemObject FSO = new FileSystemObject();
            String sPath;
            sPath = System.IO.Directory.GetParent(System.Windows.Forms.Application.StartupPath).ToString();
            //sPath = System.IO.Directory.GetParent(sPath).ToString();

            PSH_Globals.SP_XMLPath = sPath + "\\PSH_BOne_AddOn";
            PSH_Globals.SP_Path = sPath + "\\PSH_BOne_AddOn\\PathINI";

            try
            {
                Query01 = "select UniqueID from [Authority_Screen] where Gubun ='H' and updateYN ='Y'and UserID ='" + PSH_Globals.oCompany.UserName + "'";
                oRecordSet01.DoQuery(Query01);

                //���� ���� ����
                if (FSO.FolderExists(PSH_Globals.SP_XMLPath + "\\xml_temp") == false)
                {
                    FSO.CreateFolder(PSH_Globals.SP_XMLPath + "\\xml_temp");
                }
                //���� �̵�

                if (FSO.FileExists(PSH_Globals.SP_XMLPath + "\\" + PSH_Globals.oCompany.UserName + "_Menu_KOR.xml") == true)
                {
                    FSO.MoveFile(PSH_Globals.SP_XMLPath + "\\*.xml", PSH_Globals.SP_XMLPath + "\\xml_temp\\");
                }

                //������ ���� ���������� �̰�
                if (FSO.FileExists(PSH_Globals.SP_XMLPath + "\\xml_temp\\" + PSH_Globals.oCompany.UserName + "_Menu_KOR.xml") == true)
                {
                    FSO.MoveFile(PSH_Globals.SP_XMLPath + "\\xml_temp\\" + PSH_Globals.oCompany.UserName + "_Menu_KOR.xml", PSH_Globals.SP_XMLPath + "\\");
                }

                //�̰����� �� ���� ����
                FSO.DeleteFile(PSH_Globals.SP_XMLPath + "\\xml_temp\\*.*");

                if (FSO.FileExists(PSH_Globals.SP_XMLPath + "\\" + PSH_Globals.oCompany.UserName + "_Menu_KOR.xml") == false)
                {
                    SaveMenuXml();
                    //XML ����
                }

                if ((oRecordSet01.RecordCount) != 0)
                {
                    FSO.DeleteFile(PSH_Globals.SP_XMLPath + "\\" + PSH_Globals.oCompany.UserName + "_Menu_KOR.xml");
                    SaveMenuXml();
                    //XML ����
                }
                //XML No ����
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("XmlCreateYN_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// ���� �޴��� XML Client PC�� ����
        /// </summary>
        private void SaveMenuXml()
        {
            MSXML2.DOMDocument30 objDOM = new MSXML2.DOMDocument30();
            string Query01 = string.Empty;
            string UpdateQry01 = string.Empty;
            int i = 0;
            int j = 0;
            string NowType = string.Empty;
            string UserID = string.Empty;

            string AfType = string.Empty;
            string NowLevel = string.Empty;
            string AfLevel = string.Empty;

            string NowSeq = string.Empty;
            string AfSeq = string.Empty;

            string teststring = string.Empty;
            string XmlString = string.Empty;

            string oFilePath = string.Empty;
            MSXML2.DOMDocument xmldoc = new MSXML2.DOMDocument();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                AfLevel = "0";
                NowLevel = "0";
                UserID = PSH_Globals.oCompany.UserName;

                Query01 = "exec [PS_SY004_01] '" + UserID + "','H'";
                oRecordSet01.DoQuery(Query01);

                XmlString = "<Application><Menus><action type=\"add\">";

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    NowType = oRecordSet01.Fields.Item("type").Value;
                    NowSeq = oRecordSet01.Fields.Item("Seq").Value;

                    if (NowType == "2")
                    {
                        NowLevel = oRecordSet01.Fields.Item("level").Value;
                    }

                    if (i != oRecordSet01.RecordCount - 1)
                    {
                        oRecordSet01.MoveNext();
                    }
                    else
                    {
                    }

                    AfType = oRecordSet01.Fields.Item("type").Value;
                    AfSeq = oRecordSet01.Fields.Item("Seq").Value;

                    if (AfType == "2")
                    {
                        AfLevel = oRecordSet01.Fields.Item("level").Value;
                    }

                    if (i != oRecordSet01.RecordCount - 1)
                    {
                        oRecordSet01.MovePrevious();
                    }

                    XmlString = XmlString + "<Menu Checked=\"0\" Enabled=\"1\" FatherUID=\"" + oRecordSet01.Fields.Item("FatherID").Value + "\"";
                    XmlString = XmlString + " position=\"" + oRecordSet01.Fields.Item("position").Value + "\"";

                    XmlString = XmlString + " String=\"" + oRecordSet01.Fields.Item("String").Value + "\"";
                    XmlString = XmlString + " Type=\"" + oRecordSet01.Fields.Item("type").Value + "\"";
                    XmlString = XmlString + " UniqueID=\"" + oRecordSet01.Fields.Item("UniqueID").Value + "\"";

                    if (oRecordSet01.Fields.Item("UniqueID").Value == "IFX00000000F")
                    {
                        XmlString = XmlString + " Image=\"\\\\191.1.1.220\\b1_shr\\PathINI\\QM.jpg\"";

                    }
                    else if (oRecordSet01.Fields.Item("UniqueID").Value == "HGA00000000F")
                    {
                        XmlString = XmlString + " Image=\"\\\\191.1.1.220\\b1_shr\\PathINI\\GA.jpg\"";

                    }
                    else if (oRecordSet01.Fields.Item("UniqueID").Value == "GQM00000000F")
                    {
                        XmlString = XmlString + " Image=\"\\\\191.1.1.220\\b1_shr\\PathINI\\QM.jpg\"";
                    }


                    if (NowType == "2")
                    {
                        XmlString = XmlString + ">";
                    }
                    else
                    {
                        XmlString = XmlString + "/>";
                    }

                    // �������� �ݴ� �κ�
                    if ((i == oRecordSet01.RecordCount - 1))
                    {

                        if (Convert.ToDouble(NowType) == 2 && Convert.ToDouble(NowLevel) == 1)
                        {
                            XmlString = XmlString + "</Menu>";

                            for (j = Convert.ToInt32(NowLevel) - 1; j >= 0; j += -1)
                            {
                                XmlString = XmlString + "</action></Menus></Menu>";
                            }

                        }
                        else if (Convert.ToDouble(NowType) == 1 && Convert.ToDouble(NowLevel) == 1)
                        {
                            for (j = Convert.ToInt32(NowLevel); j >= 0; j += -1)
                            {
                                XmlString = XmlString + "</action></Menus></Menu>";
                            }

                        }
                        else if (Convert.ToDouble(NowType) == 2 && Convert.ToDouble(NowLevel) == 2)
                        {
                            XmlString = XmlString + "</Menu>";

                            for (j = Convert.ToInt32(NowLevel); j >= 0; j += -1)
                            {
                                XmlString = XmlString + "</action></Menus></Menu>";
                            }

                        }
                        else
                        {
                            for (j = Convert.ToInt32(NowLevel); j >= 0; j += -1)
                            {
                                XmlString = XmlString + "</action></Menus></Menu>";
                            }
                        }
                    }
                    else
                    {
                        if ((Convert.ToDouble(NowType) == 1 && Convert.ToDouble(AfType) == 1) && (NowLevel == AfLevel) && (Strings.Left(NowSeq, 9) == Strings.Left(AfSeq, 9)))
                        {
                        }
                        else if ((Convert.ToDouble(NowType) == 1 && Convert.ToDouble(AfType) == 1) && (NowLevel == AfLevel) && (Strings.Left(NowSeq, 9) != Strings.Left(AfSeq, 9)))
                        {
                            XmlString = XmlString + "</action></Menus></Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 1 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 1 && Convert.ToDouble(AfLevel) == 0) && (Strings.Left(NowSeq, 9) != Strings.Left(AfSeq, 9)) && Strings.Right(Strings.Left(NowSeq, 5), 2) == "99")
                        {
                            XmlString = XmlString + "</action></Menus></Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 1 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 1 && Convert.ToDouble(AfLevel) == 0) && (Strings.Left(NowSeq, 9) != Strings.Left(AfSeq, 9)) && Strings.Right(Strings.Left(NowSeq, 5), 2) != "99")
                        {
                            XmlString = XmlString + "</action></Menus></Menu></action></Menus></Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 1 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 2 && Convert.ToDouble(AfLevel) == 0))
                        {
                            XmlString = XmlString + "</action></Menus></Menu></action></Menus></Menu></action></Menus></Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 1 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 0 && Convert.ToDouble(AfLevel) == 0))
                        {
                            XmlString = XmlString + "</action></Menus></Menu>";
                        }
                        else if (((Convert.ToDouble(NowType) == 2 && Convert.ToDouble(AfType) == 1) && (NowLevel == AfLevel)) && Strings.Left(NowSeq, 9) == Strings.Left(AfSeq, 9))
                        {
                            XmlString = XmlString + "<Menus><action type=\"add\">";
                        }
                        else if (((Convert.ToDouble(NowType) == 2 && Convert.ToDouble(AfType) == 1) && (NowLevel == AfLevel)) && (Strings.Left(NowSeq, 9) != Strings.Left(AfSeq, 9)))
                        {
                            XmlString = XmlString + "</Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 2 && Convert.ToDouble(AfType) == 1) && (Convert.ToDouble(NowLevel) == 0 && Convert.ToDouble(AfLevel) == 0))
                        {
                            XmlString = XmlString + "<Menus><action type=\"add\">";
                        }
                        else if ((Convert.ToDouble(NowType) == 1 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 2 && Convert.ToDouble(AfLevel) == 2))
                        {
                            XmlString = XmlString + "</action></Menus></Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 1 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 1 && Convert.ToDouble(AfLevel) == 1))
                        {
                            XmlString = XmlString + "</action></Menus></Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 1 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 2 && Convert.ToDouble(AfLevel) == 1) && (Strings.Left(NowSeq, 9) != Strings.Left(AfSeq, 9)) && Strings.Right(Strings.Left(NowSeq, 5), 2) != "99" && Strings.Right(Strings.Left(NowSeq, 7), 2) != "99")
                        {
                            XmlString = XmlString + "</action></Menus></Menu></action></Menus></Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 1 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 2 && Convert.ToDouble(AfLevel) == 1))
                        {
                            XmlString = XmlString + "</action></Menus></Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 2 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 0 && Convert.ToDouble(AfLevel) == 1))
                        {
                            XmlString = XmlString + "<Menus><action type=\"add\">";
                        }
                        else if ((Convert.ToDouble(NowType) == 2 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 1 && Convert.ToDouble(AfLevel) == 2))
                        {
                            XmlString = XmlString + "<Menus><action type=\"add\">";
                        }
                        else if ((Convert.ToDouble(NowType) == 2 && Convert.ToDouble(AfType) == 2) && (NowLevel == AfLevel))
                        {
                            XmlString = XmlString + "</Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 2 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 2 && Convert.ToDouble(AfLevel) == 1))
                        {
                            XmlString = XmlString + "</Menu></action></Menus></Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 2 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 2 && Convert.ToDouble(AfLevel) == 0))
                        {
                            XmlString = XmlString + "</Menu></action></Menus></Menu></action></Menus></Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 2 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 1 && Convert.ToDouble(AfLevel) == 0))
                        {
                            XmlString = XmlString + "</Menu></action></Menus></Menu>";
                        }
                        else
                        {
                            XmlString = XmlString + "<err>";
                        }
                    }
                    oRecordSet01.MoveNext();
                }

                XmlString = XmlString + "</action></Menus></Application>";

                xmldoc.loadXML(XmlString);
                xmldoc.insertBefore(xmldoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-16'"), xmldoc.childNodes[0]);

                oFilePath = PSH_Globals.SP_XMLPath + "\\";

                UserID = UserID + "_Menu_KOR.xml";
                xmldoc.save(oFilePath + UserID);

                UpdateQry01 = "update [Authority_Screen] set UpdateYN ='N' where Gubun ='H' and updateYN ='Y'and UserID ='" + PSH_Globals.oCompany.UserName + "'";
                oRecordSet01.DoQuery(UpdateQry01);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("SaveMenuXml_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xmldoc);
            }
        }

        /// <summary>
        /// ���� �޴��� XML �ε�
        /// </summary>
        private void Load_MenuXml()
        {
            string FileName = string.Empty;
            System.Xml.XmlDocument oXmlDoc = null;
            oXmlDoc = new System.Xml.XmlDocument();

            FileName = PSH_Globals.oCompany.UserName + "_Menu_KOR.xml";

            oXmlDoc.Load(PSH_Globals.SP_XMLPath + "\\" + FileName);

            string tmpStr;
            tmpStr = oXmlDoc.InnerXml;
            PSH_Globals.SBO_Application.LoadBatchActions(tmpStr);
            //sPath = PSH_Globals.SBO_Application.GetLastBatchResults();
        }

        /// <summary>
        /// ��ȿ�� ������ �˻�
        /// </summary>
        /// <param name="FormType"></param>
        /// <returns></returns>
        private bool Check_ValidateForm(string FormType)
        {
            bool functionReturnValue = false;

            try
            {
                for (int i = 1; i <= PSH_Globals.FormTypeListCount; i++)
                {
                    if (PSH_Globals.FormTypeList[i].ToString() == FormType)
                    {
                        functionReturnValue = true;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Check_ValidateForm_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// AddOn �߰� �� ����
        /// </summary>
        /// <param name="pVal"></param>
        /// <param name="pBaseClass"></param>
		private void Create_USERForm(SAPbouiCOM.MenuEvent pVal, ref PSH_BaseClass pBaseClass)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        #region �λ� ����
                        case "PH_PY001": //��������͵��

                            pBaseClass = new PH_PY001();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY005": //������������  

                            pBaseClass = new PH_PY005();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY006": //��ȣ�۾����

                            pBaseClass = new PH_PY006();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY011": //������ȣĪ�ϰ�����

                            pBaseClass = new PH_PY011();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY017": //����������

                            pBaseClass = new PH_PY017();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY021": //�����󿬶�ó����

                            pBaseClass = new PH_PY021();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY201": //�����ӹ��� �ް���� ���

                            pBaseClass = new PH_PY201();
                            pBaseClass.LoadForm("");
                            break;


                        case "PH_PY204": //������ȹ���

                            pBaseClass = new PH_PY204();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY203": //�����������

                            pBaseClass = new PH_PY203();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY205": //������ȹVS������ȸ

                            pBaseClass = new PH_PY205();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY009": //����������UPLOAD

                            pBaseClass = new PH_PY009();
                            pBaseClass.LoadForm("");
                            break;
						
						case "PH_PY202": //�����ӹ��� �ް���� ��� ��Ȳ

                            pBaseClass = new PH_PY202();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY510": //������

                            pBaseClass = new PH_PY510();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY522": //�ӱ���ũ ����� ��Ȳ

                            pBaseClass = new PH_PY522();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY523": //�ӱ���ũ ����ڿ��� ������Ȳ

                            pBaseClass = new PH_PY523();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY524": //������ �߰����곻��

                            pBaseClass = new PH_PY524();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY419": //ǥ�ؼ����������ڵ��

                            pBaseClass = new PH_PY419();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY016": //�⺻�������

                            pBaseClass = new PH_PY016();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY775": //���κ� ������Ȳ

                            pBaseClass = new PH_PY775();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY776": //�ܿ�������Ȳ(������)

                            pBaseClass = new PH_PY776();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA65": //������Ȳ(����)

                            pBaseClass = new PH_PYA65();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY570": //����/���ϱٹ���Ȳ

                            pBaseClass = new PH_PY570();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY585": //������ٱ�Ϻ�

                            pBaseClass = new PH_PY585();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY610": //������ٱ�Ϻ�

                            pBaseClass = new PH_PY610();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY615": //�����ٹ���Ȳ
                            pBaseClass = new PH_PY615();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY620": //���������ϱٹ�����Ȳ

                            pBaseClass = new PH_PY620();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY675": //�ٹ�����Ȳ

                            pBaseClass = new PH_PY675();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA60": //���ڱݽ�û����(����)

                            pBaseClass = new PH_PYA60();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY625": //��Ź�ڸ��

                            pBaseClass = new PH_PY625();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY630": //��Ź�ڸ��

                            pBaseClass = new PH_PY630();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY700": //�޿����޴���

                            pBaseClass = new PH_PY700();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY301": //���ڱݽ�û���

                            pBaseClass = new PH_PY301();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY701": //�޿����޴���(������)

                            pBaseClass = new PH_PY701();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA20": //�޿��μ����������(�μ�)

                            pBaseClass = new PH_PYA20();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA40": //�󿩺μ����������(�μ�)

                            pBaseClass = new PH_PYA40();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA50": //DC��ȯ�ںδ�����޳���

                            pBaseClass = new PH_PYA50();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA75": //�����ܼ������޴���

                            pBaseClass = new PH_PYA75();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY765": //�޿�����������

                            pBaseClass = new PH_PY765();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY680": //�����Ȳ

                            pBaseClass = new PH_PY680();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY860": //ȣ��ǥ��ȸ

                            pBaseClass = new PH_PY860();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY508": //�������� ��� �� �߱�

                            pBaseClass = new PH_PY508();
                            pBaseClass.LoadForm("");
                            break;
                            
                        case "PH_PY770": //�����ҵ��õ¡�����������

                            pBaseClass = new PH_PY770();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY780": //����뺸�賻��

                            pBaseClass = new PH_PY780();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY785": //�����ο��ݳ���

                            pBaseClass = new PH_PY785();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY790": //���ǰ����賻��

                            pBaseClass = new PH_PY790();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY795": //�����μ����޿���Ȳ  

                            pBaseClass = new PH_PY795();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY805": //�޿����纯������

                            pBaseClass = new PH_PY805();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY810": //���޺�����ӱݳ���

                            pBaseClass = new PH_PY810();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY815": //��ü����ӱݳ��� 

                            pBaseClass = new PH_PY815();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY820": //����ӱݳ���

                            pBaseClass = new PH_PY820();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY825": // ������O/T��Ȳ

                            pBaseClass = new PH_PY825();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY830": // �μ����ΰǺ���Ȳ(��ȹ)

                            pBaseClass = new PH_PY830();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY835": // ���޺�O/T�׼�����Ȳ

                            pBaseClass = new PH_PY835();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY840": // ǳ�����ڰ����ڷ�

                            pBaseClass = new PH_PY840();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY845": // �Ⱓ���޿����޳���

                            pBaseClass = new PH_PY845();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY850": // �ұ޺����޸���

                            pBaseClass = new PH_PY850();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY855": // ���κ��ӱ����޴���

                            pBaseClass = new PH_PY855();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY865": // ��뺸����Ȳ(����)
                            pBaseClass = new PH_PY865();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY870": // ��纰��O/T�׼�����Ȳ   
                            pBaseClass = new PH_PY870();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY875": // ���޺������������  
                            pBaseClass = new PH_PY875();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY716": // �Ⱓ���޿��μ����������
                            pBaseClass = new PH_PY716();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY721": // �Ⱓ���󿩺μ����������
                            pBaseClass = new PH_PY721();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY717": // �޿��ݺ��������(��ȹ��)
                            pBaseClass = new PH_PY717();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY715": // �޿��μ����������
                            pBaseClass = new PH_PY715();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY720": // �󿩺μ����������
                            pBaseClass = new PH_PY720();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY725": // �޿����޺�������� 
                            pBaseClass = new PH_PY725();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY740": // �����޺��������
                            pBaseClass = new PH_PY740();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY745": // �������޴���   
                            pBaseClass = new PH_PY745();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY750": // �ٷμҵ�¡����Ȳ
                            pBaseClass = new PH_PY750();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY755": // ��ȣȸ������Ȳ
                            pBaseClass = new PH_PY755();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY760": // ����ӱݹ������ݻ��⳻����
                            pBaseClass = new PH_PY760();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY019": // �ݺ�����

                            pBaseClass = new PH_PY019();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY018": // ���������ϱ����üũ

                            pBaseClass = new PH_PY018();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY117": // �޻󿩸����۾�

                            pBaseClass = new PH_PY117();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY123": // ���з����

                            pBaseClass = new PH_PY123();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY409": // ��α��������ڷ���

                            pBaseClass = new PH_PY409();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY555": // ���ϱٹ�����Ȳ

                            pBaseClass = new PH_PY555();
                            pBaseClass.LoadForm("");
                            break;

						case "PH_PY010": //���ϱ���ó��

                            pBaseClass = new PH_PY010();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY013": //�����ڵ���(���)

                            pBaseClass = new PH_PY013();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY014": //�����ϼ�����

                            pBaseClass = new PH_PY014();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY583": //���¸���üũ

                            pBaseClass = new PH_PY583();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY120": //�ұ޺б޿�����

                            pBaseClass = new PH_PY120();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY133": //������ Ƚ�� ����

                            pBaseClass = new PH_PY133();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY119": //�޻��������ϻ���

                            pBaseClass = new PH_PY119();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY002": //���½ð����е��

                            pBaseClass = new PH_PY002();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY101": //����� ���

                            pBaseClass = new PH_PY101();
                            pBaseClass.LoadForm("");
                            break;
                            
                         case "PH_PY134": //�ҵ漼 / �ֹμ� ����

                            pBaseClass = new PH_PY134();
                            pBaseClass.LoadForm("");
                            break;
                            
                        case "PH_PY100": //���ؼ��׼���

                            pBaseClass = new PH_PY100();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY109_1": //�޻󿩺����ڷ��׸����

                            pBaseClass = new PH_PY109_1();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY131": //�޻󿩺����ڷ��׸����

                            pBaseClass = new PH_PY131();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY129": //���κ� ��������(DC��) ���

                            pBaseClass = new PH_PY129();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY314": //��αݰ�� ���� ��ȸ(�޿������ڷ��)

                            pBaseClass = new PH_PY314();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY695": //�λ���ī��

                            pBaseClass = new PH_PY695();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY605": //�ټӺ����ް��߻��׻�볻��

                            pBaseClass = new PH_PY605();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY560": //�������Ȳ

                            pBaseClass = new PH_PY560();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY565": //����ٹ�����Ȳ

                            pBaseClass = new PH_PY565();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY575": //���±�����Ȳ

                            pBaseClass = new PH_PY575();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY580": //���κ����¿���

                            pBaseClass = new PH_PY580();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY985": //��αݰ�� ���� ��ȸ(�޿������ڷ��)

                            pBaseClass = new PH_PY985();
                            pBaseClass.LoadForm();
                            break;


                        case "PH_PY590": //�Ⱓ����������ǥ

                            pBaseClass = new PH_PY590();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY600": //���ں�����ٹ���Ȳ

                            pBaseClass = new PH_PY600();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY705": //��������ޱ���Ȯ��
                            pBaseClass = new PH_PY705();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY676": //���½ð�������ȸ
                            pBaseClass = new PH_PY676();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY679": //���κ������ڷ�����
                            pBaseClass = new PH_PY679();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY681": //��ٹ��ϼ���Ȳ
                            pBaseClass = new PH_PY681();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY645": //�ڰݼ���������Ȳ
                            pBaseClass = new PH_PY645();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA55": //����¡����ȯ�޴���(����)
                            pBaseClass = new PH_PYA55();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY550": //��ü�ο���Ȳ
                            pBaseClass = new PH_PY550();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY650": //�뵿���հ�����Ȳ
                            pBaseClass = new PH_PY650();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY685": //���󰡱���Ȳ
                            pBaseClass = new PH_PY685();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY690": //��������Ȳ
                            pBaseClass = new PH_PY690();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA70": //�ҵ漼��õ¡������������û�����
                            pBaseClass = new PH_PYA70();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY501": //���ǹ߱���Ȳ
                            pBaseClass = new PH_PY501();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY525": //�зº��ο���Ȳ
                            pBaseClass = new PH_PY525();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY935": //�����ȣ��Ȳ
                            pBaseClass = new PH_PY935();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY530": //���ɺ��ο���Ȳ
                            pBaseClass = new PH_PY530();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY505": //�Ի��ڴ���
                            pBaseClass = new PH_PY505();
                            pBaseClass.LoadForm(""); 
                            break;

                        case "PH_PY520": //���������������ڴ���
                            pBaseClass = new PH_PY520(); 
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY640": //���ο���������ȯ����Ȳ
                            pBaseClass = new PH_PY640();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY545": //�ο���Ȳ(�볻��)
                            pBaseClass = new PH_PY545();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY655": //���ƴ������Ȳ
                            pBaseClass = new PH_PY655();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY660": //��ֱٷ�����Ȳ
                            pBaseClass = new PH_PY660(); 
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY540": //�ο���Ȳ(��ܿ�)
                            pBaseClass = new PH_PY540();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY551": //����ο���ȸ
                            pBaseClass = new PH_PY551();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY535": //�ټӳ�����ο���Ȳ
                            pBaseClass = new PH_PY535();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY665": //����ڳ���Ȳ
                            pBaseClass = new PH_PY665();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY670": //���κ�������Ȳ
                            pBaseClass = new PH_PY670();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY503": //��������ڸ��
                            pBaseClass = new PH_PY503();
                            pBaseClass.LoadForm(""); 
                            break;

                        case "PH_PY507": //��������Ȳ
                            pBaseClass = new PH_PY507(); 
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY635": //����,��������Ȳ
                            pBaseClass = new PH_PY635();
                            pBaseClass.LoadForm(""); 
                            break;

                        case "PH_PY515": //�����ڻ�����
                            pBaseClass = new PH_PY515();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY003": //���¿��µ��

                            pBaseClass = new PH_PY003();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY595": //�ټӳ����Ȳ
                            pBaseClass = new PH_PY595();
                            pBaseClass.LoadForm(""); 
                            break;

                        case "PH_PY931": //ǥ�ؼ�������������ȸ
                            pBaseClass = new PH_PY931();
                            pBaseClass.LoadForm(""); 
                            break;

                        case "PH_PY932": //���ٹ��������Ȳ
                            pBaseClass = new PH_PY932();
                            pBaseClass.LoadForm(""); 
                            break;

                        case "PH_PY933": //�����Ѿ׽Ű�����ڷ�
                            pBaseClass = new PH_PY933();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY800": //�ΰǺ������ڷ�
                            pBaseClass = new PH_PY800();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY004": //�ٹ��������

                            pBaseClass = new PH_PY004();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA30": //�����޴���(�μ�)
                            pBaseClass = new PH_PYA30();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA10": //�޿����޴���(�μ�)
                            pBaseClass = new PH_PYA10();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY915": //�ٷμҵ��õ¡�������
                            pBaseClass = new PH_PY915();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY735": //�󿩺������
                            pBaseClass = new PH_PY735();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY730": //�޿��������
                            pBaseClass = new PH_PY730();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY710": //�����޴���
                            pBaseClass = new PH_PY710();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY930": //����¡����ȯ�޴���
                            pBaseClass = new PH_PY930();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY925": //��αݸ������
                            pBaseClass = new PH_PY925();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY718": //����Ϸ�ݾ״��O/T��Ȳ
                            pBaseClass = new PH_PY718();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY920": //��õ¡�����������
                            pBaseClass = new PH_PY920();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY311": //��ٹ���������

                            pBaseClass = new PH_PY311();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY910": //�ҵ�����Ű����
                            pBaseClass = new PH_PY910();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY401": //���ٹ������
                            pBaseClass = new PH_PY401();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY402": //������ʵ��
                            pBaseClass = new PH_PY402();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY405": //�Ƿ���ڷ���
                            pBaseClass = new PH_PY405();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY407": //�����αݵ��
                            pBaseClass = new PH_PY407();
                            pBaseClass.LoadForm(""); 
                            break;

                        case "PH_PY411": //���������ҵ�������
                            pBaseClass = new PH_PY411();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY413": //������.�����������Ա��ڷ� ���
                            pBaseClass = new PH_PY413();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY113": //��(��)�� �а��� ����
                            pBaseClass = new PH_PY113();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY980": //�ٷμҵ����޸����ڷ� �����ü����
                            pBaseClass = new PH_PY980();
                            pBaseClass.LoadForm();
                            break;

                        case "PH_PY995": //�����ҵ����޸����ڷ� �����ü����
                            pBaseClass = new PH_PY995();
                            pBaseClass.LoadForm();
                            break;

                        case "PH_PY677": //���ϱ����̻�����ȸ
                            pBaseClass = new PH_PY677();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY678": //�����ٹ����ϰ����
                            pBaseClass = new PH_PY678();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY683": //����ٹ�������Ȳ
                            pBaseClass = new PH_PY683();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY307": //���ڱݽ�û����(�б⺰)
                            pBaseClass = new PH_PY307();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY306": //���ڱݽ�û����(���κ�)
                            pBaseClass = new PH_PY306();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY305": //���ڱݽ�û��
                            pBaseClass = new PH_PY305();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY315": //���κ� ��α� �ܾ���Ȳ
                            pBaseClass = new PH_PY315();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY313": //��αݰ��
                            pBaseClass = new PH_PY313();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY107": //�޻󿩱����ϼ���
                            pBaseClass = new PH_PY107();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY122": //�޻����_���κμ��������
                            pBaseClass = new PH_PY122();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY015": //�����̿����
                            pBaseClass = new PH_PY015();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY111": //�޻󿩰��
                            pBaseClass = new PH_PY111();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY103": //�����׸���
                            pBaseClass = new PH_PY103();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY312": //���κ�������ݵ��(â��)
                            pBaseClass = new PH_PY312();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY121": //���κ� �򰡰��޾� ���
                            pBaseClass = new PH_PY121();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY415": //�ҵ�������
                            pBaseClass = new PH_PY415();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY302": //���ڱ����޿Ϸ�ó��
                            pBaseClass = new PH_PY302();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY303": //���ڱ��������ϻ���
                            pBaseClass = new PH_PY303();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY417": //���ڱ��������ϻ���
                            pBaseClass = new PH_PY417();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY007": //�����ܰ����
                            pBaseClass = new PH_PY007();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY109": //�޻󿩺����ڷ���
                            pBaseClass = new PH_PY109();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY000": //��������͵��
                            pBaseClass = new PH_PY000();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY008": //�ϱ��µ��
                            pBaseClass = new PH_PY008();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY030": //������
                            pBaseClass = new PH_PY030();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY031": //������
                            pBaseClass = new PH_PY031();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY020": // ���±⺻���� ������(N.G.Y)_�������
                            pBaseClass = new PH_PY020();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY104": // ������������ݾ��ϰ����
                            pBaseClass = new PH_PY104();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY132": // ������ ���� ���κ� ���
                            pBaseClass = new PH_PY132();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY116": // �����ݺа���
                            pBaseClass = new PH_PY116();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY105": // ȣ��ǥ���
                            pBaseClass = new PH_PY105();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY108": //�����޷�����
                            pBaseClass = new PH_PY108();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY125": //���κ� �������� �������(���� Upload)
                            pBaseClass = new PH_PY125();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY102": //�����׸���
                            pBaseClass = new PH_PY102();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY130": //���� ���������� ��޵��
                            pBaseClass = new PH_PY130();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY032": //��������
                            pBaseClass = new PH_PY032();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY110": //���κ��������
                            pBaseClass = new PH_PY110();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY106": //������ļ���
                            pBaseClass = new PH_PY106();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY114": //������ļ���
                            pBaseClass = new PH_PY114();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY112": //������ļ���
                            pBaseClass = new PH_PY112();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY309": //��αݵ��
                            pBaseClass = new PH_PY309();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY034": //����а�ó��
                            pBaseClass = new PH_PY034();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY124": //�����Ǿ� �ݾ׵��
                            pBaseClass = new PH_PY124();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY999": //�޴����Ѱ���
                            pBaseClass = new PH_PY999();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA80": //�ٹ��ð�ǥ���
                            pBaseClass = new PH_PYA80();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA90": //�ٷμҵ氣�����޸���(�������Ű����ϻ���)
                            pBaseClass = new PH_PYA90();
                            pBaseClass.LoadForm();
                            break;

                        case "PH_PY526": //�ӱ���ũ�ο���Ȳ
                            pBaseClass = new PH_PY526();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY127": //���κ� 4�뺸�� �������� �� ����ݾ� ���(���� Upload)
                            pBaseClass = new PH_PY127();
                            pBaseClass.LoadForm("");
                            break;
                            
                        case "PH_PY310": //��αݰ�����ȯ(2019.11.21 �۸��)
    						pBaseClass = new PH_PY310();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY115": //�����ݰ��(2019.11.22 �۸��)
							pBaseClass = new PH_PY115();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY118": //�޻� E-Mail �߼�(2019.12.16 �۸��)
							pBaseClass = new PH_PY118();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY135": //�޻󿩺а�ó��(2019.12.30 �۸��)
                            pBaseClass = new PH_PY135();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY136": //�޻󿩺а�ó�� ��α�Ģ����(2020.02.06 �۸��)
                            pBaseClass = new PH_PY136();
                            pBaseClass.LoadForm("");
                            break;
                        #endregion ����

                        #region � ����
                        case "PS_DateChange": //��¥����ó��
                            pBaseClass = new PS_DateChange();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_DateCommit": //��¥�������
                            pBaseClass = new PS_DateCommit();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY998": //����� ���� ��ȸ
                            pBaseClass = new PH_PY998();
                            pBaseClass.LoadForm("");
                            break;
                        #endregion

                        #region �繫 ����
                        case "PS_CO020": //������ұ׷���
                            pBaseClass = new PS_CO020();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO010": //������ҵ��
                            pBaseClass = new PS_CO010();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO001": //��긶������
                            pBaseClass = new PS_CO001();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO658": //�����繫��ǥ ���� ����
                            pBaseClass = new PS_CO658();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO600": //�����繫��ǥ
                            pBaseClass = new PS_CO600();
							pBaseClass.LoadForm("");
                            break;

                        case "PS_CO605": //�����繫��ǥ
                            pBaseClass = new PS_CO605();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO610": //�����ڻ� ������ ��ü
                            pBaseClass = new PS_CO610();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO660": //�Ⱓ�����
                            pBaseClass = new PS_CO660();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO665": //�Ⱓ�����Ȳ(����)
                            pBaseClass = new PS_CO665();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO670": //�Ⱓ���а����
                            pBaseClass = new PS_CO670();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO120": //������ �������
                            pBaseClass = new PS_CO120();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO130": //��ǰ�� �������
                            pBaseClass = new PS_CO130();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO131": //������������Ȳ
                            pBaseClass = new PS_CO131();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO510": //����������������ȸ
                            pBaseClass = new PS_CO510();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO520": //��ǰ���� �����׸� ��ȸ
                            pBaseClass = new PS_CO520();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO501": //ǰ�񺰿������
                            pBaseClass = new PS_CO501();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO502": //ǰ����տ����׸���
                            pBaseClass = new PS_CO502();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO503": //���ϰ���׹׻���������
                            pBaseClass = new PS_CO503();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO504": //�����ǸŹ׻�������
                            pBaseClass = new PS_CO504();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO210": //5.������ǰ�������
                            pBaseClass = new PS_CO210();
                            pBaseClass.LoadForm("");
                            break;
                            #endregion
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Create_USERForm_Error: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// �ý����� ����
        /// </summary>
        /// <param name="pval"></param>
        private void Create_SYSTEMForm(SAPbouiCOM.ItemEvent pval)
        {
            try
            {
                if (pval.BeforeAction == true)
                {
                    if (pval.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    {
                        //switch (pval.FormTypeEx)
                        //{

                            //Case "-60100"       '//�λ����>��������͵����� (����� ���� �ʵ�)
                            //  Set oTempClass = New SM60100: oTempClass.LoadForm (pval.FormUID): AddForms oTempClass, pval.FormUID, pval.FormTypeEx
                        //}
                    }
                }
                return;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("Create_SYSTEMForm_Error: " + ex.Message, BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// ������ �� Ŭ���� ����(������:exe �޸� ���� �ȵ�, 2018.12.03 �۸��)
        /// </summary>
        private void TerminateApplication()
        {
            int i = 0;
            PSH_BaseClass oTempClass = new PSH_BaseClass();

            if (PSH_Globals.ClassList.Count > 0)
            {
                for (i = 0; i <= PSH_Globals.ClassList.Count - 1; i++)
                {
                    oTempClass = (PSH_BaseClass)PSH_Globals.ClassList[i];
                    PSH_Globals.ClassList.Remove(i);
                }
            }

            PSH_Globals.oCompany.Disconnect();
        }

        #region �̺�Ʈ

        private void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    //this.TerminateApplication();
                    PSH_Globals.oCompany.Disconnect();
                    System.Environment.Exit(0);
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    Load_MenuXml();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    //this.TerminateApplication();
                    PSH_Globals.oCompany.Disconnect();
                    System.Environment.Exit(0);
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //this.TerminateApplication();
                    PSH_Globals.oCompany.Disconnect();
                    System.Environment.Exit(0);
                    break;
            }
        }

        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            PSH_BaseClass oTempClass = new PSH_BaseClass();
            string FormUID = string.Empty;

            try
            {
                if (pVal.BeforeAction == true)
                {
                    Create_USERForm(pVal, ref oTempClass);
                }

                FormUID = PSH_Globals.SBO_Application.Forms.ActiveForm.UniqueID;

                if (Strings.Left(FormUID, 2) != "F_")
                {
                    if (Check_ValidateForm(PSH_Globals.SBO_Application.Forms.ActiveForm.TypeEx))
                    {
                        oTempClass = (PSH_BaseClass)PSH_Globals.ClassList[FormUID];
                        if (oTempClass.oForm == null)
                        {
                            return;
                        }
                        else
                        {
                            oTempClass.Raise_FormMenuEvent(FormUID, ref pVal, ref BubbleEvent);
                        }
                    }
                }
                else if (Strings.Left(FormUID, 2) == "F_")
                {
                    if (Check_ValidateForm(PSH_Globals.SBO_Application.Forms.ActiveForm.TypeEx))
                    {
                        oTempClass = (PSH_BaseClass)PSH_Globals.ClassList[FormUID];
                        oTempClass.Raise_FormMenuEvent(FormUID, ref pVal, ref BubbleEvent);
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("SBO_Application_MenuEvent_Error: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            PSH_BaseClass oTempClass = new PSH_BaseClass();

            try
            {
                Create_SYSTEMForm(pVal);

                if (Strings.Left(pVal.FormUID, 2) != "F_")
                {
                    if (Check_ValidateForm(pVal.FormTypeEx))
                    {
                        if (pVal.EventType == BoEventTypes.et_FORM_UNLOAD)
                        {
                            if (pVal.Before_Action == true)
                            {
                                oTempClass = (PSH_BaseClass)PSH_Globals.ClassList[FormUID];
                            }
                            else if (pVal.Before_Action == false) //FORM_UNLOAD �̺�Ʈ�� Before_Action == false �� ���� PSH_Globals.ClassList[FormUID] �� index ���� �߻��ϹǷ� ���� return
                            {
                                return;
                            }
                        }
                        else
                        {
                            oTempClass = (PSH_BaseClass)PSH_Globals.ClassList[FormUID];
                        }

                        if (oTempClass.oForm == null)
                        {
                            return;
                        }
                        else
                        {
                            oTempClass.Raise_FormItemEvent(FormUID, ref pVal, ref BubbleEvent);
                        }
                    }
                }
                else if (Strings.Left(pVal.FormUID, 2) == "F_")
                {
                    if (Check_ValidateForm(pVal.FormTypeEx))
                    {
                        oTempClass = (PSH_BaseClass)PSH_Globals.ClassList[FormUID];
                        oTempClass.Raise_FormItemEvent(FormUID, ref pVal, ref BubbleEvent);
                    }
                }
                return;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("SBO_Application_ItemEvent_Error : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        private void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            PSH_BaseClass oTempClass = new PSH_BaseClass();
            string FormUID = string.Empty;

            try
            {
                FormUID = BusinessObjectInfo.FormUID;

                if (Strings.Left(FormUID, 2) != "F_")
                {
                    if (Check_ValidateForm(BusinessObjectInfo.FormTypeEx))
                    {
                        oTempClass = (PSH_BaseClass)PSH_Globals.ClassList[FormUID];
                        if (oTempClass.oForm == null)
                        {
                            return;
                        }
                        else
                        {
                            oTempClass.Raise_FormDataEvent(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        }
                    }
                }
                else if (Strings.Left(FormUID, 2) == "F_")
                {
                    if (Check_ValidateForm(BusinessObjectInfo.FormTypeEx))
                    {
                        oTempClass = (PSH_BaseClass)PSH_Globals.ClassList[FormUID];
                        oTempClass.Raise_FormDataEvent(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                    }
                }
                return;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("SBO_Application_FormDataEvent_Error : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        private void SBO_Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            PSH_BaseClass oTempClass = new PSH_BaseClass();
            string FormUID = string.Empty;

            try
            {
                FormUID = eventInfo.FormUID;

                if (Strings.Left(FormUID, 2) != "F_")
                {
                    if (Check_ValidateForm(PSH_Globals.SBO_Application.Forms.Item(eventInfo.FormUID).TypeEx))
                    {
                        oTempClass = (PSH_BaseClass)PSH_Globals.ClassList[FormUID];

                        if (oTempClass.oForm == null)
                        {
                            return;
                        }
                        else
                        {
                            oTempClass.Raise_RightClickEvent(FormUID, ref eventInfo, ref BubbleEvent);
                        }
                    }
                }
                else if (Strings.Left(FormUID, 2) == "F_")
                {
                    if (Check_ValidateForm(PSH_Globals.SBO_Application.Forms.Item(eventInfo.FormUID).TypeEx))
                    {
                        oTempClass = (PSH_BaseClass)PSH_Globals.ClassList[FormUID];
                        oTempClass.Raise_RightClickEvent(FormUID, ref eventInfo, ref BubbleEvent);
                    }
                }
                return;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("SBO_Application_RightClickEvent_Error : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        #endregion
    }
}




