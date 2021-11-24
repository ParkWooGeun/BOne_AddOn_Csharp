using Microsoft.VisualBasic;
using System;
using System.Windows.Forms;
using SAPbouiCOM;
using Scripting;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// MainClass : 광역변수 초기화, 어플리케이션 연결, DI API 연결, 회사 DB 연결, ODBC 연결용 변수 초기화, MainMenu용 XML 로딩, 유효폼 검사, AddOn 폼 생성, System 폼 생성, 이벤트 정의, 이벤트 필터 실행
    /// ZZMDC 클래스와 매칭
    /// </summary>
    internal class PSH_MainClass
    {
        /// <summary>
        /// 생성자
        /// </summary>
        public PSH_MainClass() : base()
        {
            this.Initialize_Calss(); //클래스 초기화

            //이벤트 정의
            PSH_Globals.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            PSH_Globals.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            PSH_Globals.SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
            PSH_Globals.SBO_Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);
            PSH_Globals.SBO_Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);
        }

        /// <summary>
        /// 클래스 초기화
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
                    PSH_Globals.SBO_Application.MessageBox("DI API 연결실패", 1, "Ok", "", "");
                    System.Environment.Exit(0);
                }

                // Connect To The Company Data Base
                if (!(Connect_CompanyDB() == 0))
                {
                    PSH_Globals.SBO_Application.MessageBox("회사 DB 연결실패", 1, "Ok", "", "");
                    System.Environment.Exit(0);
                }

                PSH_SetFilter.Execute(); //Event Filter Execute

                this.XmlCreateYN();
                this.Load_MenuXml();
                //DoSomething();

                Initialize_ODBC_Variable();

                PSH_Globals.SBO_Application.StatusBar.SetText("PSH_BOne_AddOn 초기화 완료", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Initialize_Calss_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 광역변수 초기화
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
        /// 어플리케이션 연결
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PSH_BOne_AddOn 시작 중...", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("PSH_BOne_AddOn 접속 실패 : " + ex.Message, "SAP Business One", MessageBoxButtons.YesNo);
            }
        }

        /// <summary>
        /// DI API 연결
        /// </summary>
        /// <returns>0 : 성공</returns>
        private int Connect_DIAPI()
        {
            int setConnectionContextReturn;

            string sCookie;
            string sConnectionContext;

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
        /// 회사 DB 연결
        /// </summary>
        /// <returns>0 : 성공</returns>
        private int Connect_CompanyDB()
        {
            int connectToCompanyReturn;

            // Establish the connection to the company database.
            connectToCompanyReturn = PSH_Globals.oCompany.Connect();

            return connectToCompanyReturn; //36,000ms ~ 40,000ms 소요
        }

        /// <summary>
        /// ODBC 연결용 변수 초기화
        /// </summary>
        private void Initialize_ODBC_Variable()
        {
            string sQry;
            string ServerName;
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            ServerName = PSH_Globals.SBO_Application.Company.ServerName;

            sQry = "  SELECT      PARAM01 AS PARAM01,";
            sQry += "             PARAM02 AS PARAM02,";
            sQry += "             PARAM03 AS PARAM03,";
            sQry += "             PARAM04 AS PARAM04,";
            sQry += "             PARAM05 AS PARAM05,";
            sQry += "             PARAM06 AS PARAM06,";
            sQry += "             PARAM07 AS PARAM07,";
            sQry += "             PARAM08 AS PARAM08";
            sQry += " FROM        PROFILE ";
            sQry += " WHERE       TYPE = 'SERVERINFO'";

            oRecordSet.DoQuery(sQry);

            if (oRecordSet.RecordCount > 0)
            {
                //ODBC
                //PSH_Globals.SP_ODBC_YN = Trim(oRecordset.Fields("Value01").Value)
                if (codeHelpClass.Right(ServerName, 3) == "223")
                {
                    PSH_Globals.SP_ODBC_Name = "MDCERP";
                }
                else
                {
                    PSH_Globals.SP_ODBC_Name = "PSHERP_TEST"; // 191.1.1.223으로 접속시 왼쪽  ODBC로 접속
                }
                PSH_Globals.SP_ODBC_IP = ServerName;
                //접속한 서버주소를 바로 가져올수 있게 기존 PARAM01에서 가져온 값을 PSH_Globals.SBO_Application.Company.ServerName
                //PSH_Globals.SP_ODBC_IP = oRecordSet.Fields.Item("PARAM01").Value.ToString().Replace("\\", "").Trim();
                PSH_Globals.SP_ODBC_DBName = PSH_Globals.oCompany.CompanyDB;
                PSH_Globals.SP_ODBC_ID = oRecordSet.Fields.Item("PARAM07").Value.ToString().Trim();
                PSH_Globals.SP_ODBC_PW = oRecordSet.Fields.Item("PARAM08").Value.ToString().Trim();
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
        }

        /// <summary>
        /// 메인 메뉴용 XML 로딩
        /// </summary>
        private void XmlCreateYN()
        {
            string Query01;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            FileSystemObject FSO = new FileSystemObject();
            String sPath;
            sPath = System.IO.Directory.GetParent(System.Windows.Forms.Application.StartupPath).ToString();
            //sPath = System.IO.Directory.GetParent(sPath).ToString();

            PSH_Globals.SP_XMLPath = sPath + "\\PSH_BOne_AddOn";
            PSH_Globals.SP_Path = sPath + "\\PSH_BOne_AddOn\\PathINI";

            try
            {
                Query01 = "select a.UniqueID from [Authority_Screen] a inner join [Authority_User] b on a.Seq = b.seq where  a.Gubun ='H' and  b.updateYN ='Y' and b.UserID ='" + PSH_Globals.oCompany.UserName + "'";
                oRecordSet01.DoQuery(Query01);

                //파일 폴더 생성
                if (FSO.FolderExists(PSH_Globals.SP_XMLPath + "\\xml_temp") == false)
                {
                    FSO.CreateFolder(PSH_Globals.SP_XMLPath + "\\xml_temp");
                }
                //파일 이동

                if (FSO.FileExists(PSH_Globals.SP_XMLPath + "\\" + PSH_Globals.oCompany.UserName + "_Menu_KOR.xml") == true)
                {
                    FSO.MoveFile(PSH_Globals.SP_XMLPath + "\\*.xml", PSH_Globals.SP_XMLPath + "\\xml_temp\\");
                }

                //접속자 파일 정상폴더로 이관
                if (FSO.FileExists(PSH_Globals.SP_XMLPath + "\\xml_temp\\" + PSH_Globals.oCompany.UserName + "_Menu_KOR.xml") == true)
                {
                    FSO.MoveFile(PSH_Globals.SP_XMLPath + "\\xml_temp\\" + PSH_Globals.oCompany.UserName + "_Menu_KOR.xml", PSH_Globals.SP_XMLPath + "\\");
                }

                //이관폴더 내 파일 삭제
                FSO.DeleteFile(PSH_Globals.SP_XMLPath + "\\xml_temp\\*.*");

                if (FSO.FileExists(PSH_Globals.SP_XMLPath + "\\" + PSH_Globals.oCompany.UserName + "_Menu_KOR.xml") == false)
                {
                    SaveMenuXml();
                    //XML 생성
                }

                if ((oRecordSet01.RecordCount) != 0)
                {
                    FSO.DeleteFile(PSH_Globals.SP_XMLPath + "\\" + PSH_Globals.oCompany.UserName + "_Menu_KOR.xml");
                    SaveMenuXml();
                    //XML 생성
                }
                //XML No 생성
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("XmlCreateYN_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 메인 메뉴용 XML Client PC에 생성
        /// </summary>
        private void SaveMenuXml()
        {
            string Query01;
            string UpdateQry01;
            int i;
            int j;
            string NowType;
            string UserID;
            string AfType;
            string NowLevel;
            string AfLevel;
            string NowSeq;gg
            string XmlString;

            string oFilePath;
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
                        XmlString += " Image=\"\\\\191.1.1.220\\b1_shr\\PathINI\\QM.jpg\"";

                    }
                    else if (oRecordSet01.Fields.Item("UniqueID").Value == "HGA00000000F")
                    {
                        XmlString += " Image=\"\\\\191.1.1.220\\b1_shr\\PathINI\\GA.jpg\"";

                    }
                    else if (oRecordSet01.Fields.Item("UniqueID").Value == "GQM00000000F")
                    {
                        XmlString += " Image=\"\\\\191.1.1.220\\b1_shr\\PathINI\\QM.jpg\"";
                    }

                    if (NowType == "2")
                    {
                        XmlString += ">";
                    }
                    else
                    {
                        XmlString += "/>";
                    }

                    // 마지막에 닫는 부분
                    if (i == oRecordSet01.RecordCount - 1)
                    {

                        if (Convert.ToDouble(NowType) == 2 && Convert.ToDouble(NowLevel) == 1)
                        {
                            XmlString += "</Menu>";

                            for (j = Convert.ToInt32(NowLevel) - 1; j >= 0; j += -1)
                            {
                                XmlString += "</action></Menus></Menu>";
                            }

                        }
                        else if (Convert.ToDouble(NowType) == 1 && Convert.ToDouble(NowLevel) == 1)
                        {
                            for (j = Convert.ToInt32(NowLevel); j >= 0; j += -1)
                            {
                                XmlString += "</action></Menus></Menu>";
                            }

                        }
                        else if (Convert.ToDouble(NowType) == 2 && Convert.ToDouble(NowLevel) == 2)
                        {
                            XmlString += "</Menu>";

                            for (j = Convert.ToInt32(NowLevel); j >= 0; j += -1)
                            {
                                XmlString += "</action></Menus></Menu>";
                            }

                        }
                        else
                        {
                            for (j = Convert.ToInt32(NowLevel); j >= 0; j += -1)
                            {
                                XmlString += "</action></Menus></Menu>";
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
                            XmlString += "</action></Menus></Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 1 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 1 && Convert.ToDouble(AfLevel) == 0) && (Strings.Left(NowSeq, 9) != Strings.Left(AfSeq, 9)) && Strings.Right(Strings.Left(NowSeq, 5), 2) == "99")
                        {
                            XmlString += "</action></Menus></Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 1 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 1 && Convert.ToDouble(AfLevel) == 0) && (Strings.Left(NowSeq, 9) != Strings.Left(AfSeq, 9)) && Strings.Right(Strings.Left(NowSeq, 5), 2) != "99")
                        {
                            XmlString += "</action></Menus></Menu></action></Menus></Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 1 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 2 && Convert.ToDouble(AfLevel) == 0))
                        {
                            XmlString += "</action></Menus></Menu></action></Menus></Menu></action></Menus></Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 1 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 0 && Convert.ToDouble(AfLevel) == 0))
                        {
                            XmlString += "</action></Menus></Menu>";
                        }
                        else if (((Convert.ToDouble(NowType) == 2 && Convert.ToDouble(AfType) == 1) && (NowLevel == AfLevel)) && Strings.Left(NowSeq, 9) == Strings.Left(AfSeq, 9))
                        {
                            XmlString += "<Menus><action type=\"add\">";
                        }
                        else if (((Convert.ToDouble(NowType) == 2 && Convert.ToDouble(AfType) == 1) && (NowLevel == AfLevel)) && (Strings.Left(NowSeq, 9) != Strings.Left(AfSeq, 9)))
                        {
                            XmlString += "</Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 2 && Convert.ToDouble(AfType) == 1) && (Convert.ToDouble(NowLevel) == 0 && Convert.ToDouble(AfLevel) == 0))
                        {
                            XmlString += "<Menus><action type=\"add\">";
                        }
                        else if ((Convert.ToDouble(NowType) == 1 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 2 && Convert.ToDouble(AfLevel) == 2))
                        {
                            XmlString += "</action></Menus></Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 1 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 1 && Convert.ToDouble(AfLevel) == 1))
                        {
                            XmlString += "</action></Menus></Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 1 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 2 && Convert.ToDouble(AfLevel) == 1) && (Strings.Left(NowSeq, 9) != Strings.Left(AfSeq, 9)) && Strings.Right(Strings.Left(NowSeq, 5), 2) != "99" && Strings.Right(Strings.Left(NowSeq, 7), 2) != "99")
                        {
                            XmlString += "</action></Menus></Menu></action></Menus></Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 1 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 2 && Convert.ToDouble(AfLevel) == 1))
                        {
                            XmlString += "</action></Menus></Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 2 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 0 && Convert.ToDouble(AfLevel) == 1))
                        {
                            XmlString += "<Menus><action type=\"add\">";
                        }
                        else if ((Convert.ToDouble(NowType) == 2 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 1 && Convert.ToDouble(AfLevel) == 2))
                        {
                            XmlString += "<Menus><action type=\"add\">";
                        }
                        else if ((Convert.ToDouble(NowType) == 2 && Convert.ToDouble(AfType) == 2) && (NowLevel == AfLevel))
                        {
                            XmlString += "</Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 2 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 2 && Convert.ToDouble(AfLevel) == 1))
                        {
                            XmlString += "</Menu></action></Menus></Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 2 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 2 && Convert.ToDouble(AfLevel) == 0))
                        {
                            XmlString += "</Menu></action></Menus></Menu></action></Menus></Menu>";
                        }
                        else if ((Convert.ToDouble(NowType) == 2 && Convert.ToDouble(AfType) == 2) && (Convert.ToDouble(NowLevel) == 1 && Convert.ToDouble(AfLevel) == 0))
                        {
                            XmlString += "</Menu></action></Menus></Menu>";
                        }
                        else
                        {
                            XmlString += "<err>";
                        }
                    }
                    oRecordSet01.MoveNext();
                }

                XmlString += "</action></Menus></Application>";

                xmldoc.loadXML(XmlString);
                xmldoc.insertBefore(xmldoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-16'"), xmldoc.childNodes[0]);

                oFilePath = PSH_Globals.SP_XMLPath + "\\";

                UserID += "_Menu_KOR.xml";
                xmldoc.save(oFilePath + UserID);

                UpdateQry01 = "update b set b.UpdateYN ='N' from [Authority_Screen] a inner join [Authority_User] b on a.Seq = b.seq where  a.Gubun ='H' and  b.updateYN ='Y' and b.UserID ='" + PSH_Globals.oCompany.UserName + "'";
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
        /// 메인 메뉴용 XML 로딩
        /// </summary>
        private void Load_MenuXml()
        {
            System.Xml.XmlDocument oXmlDoc = new System.Xml.XmlDocument();
            string FileName = PSH_Globals.oCompany.UserName + "_Menu_KOR.xml";
            oXmlDoc.Load(PSH_Globals.SP_XMLPath + "\\" + FileName);
            PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.InnerXml);
        }

        /// <summary>
        /// 유효한 폼인지 검사(Collection에 저장되어 있는 Form, 이벤트를 실행시켜야 하는 Form)
        /// </summary>
        /// <param name="FormUID"></param>returnValue
        /// <returns></returns>
        private bool Check_ValidateForm(string FormUID)
        {
            bool returnValue = false;

            try
            {
                for (int i = 1; i <= PSH_Globals.ClassList.Count; i++)
                {
                    if (PSH_Globals.ClassList.Contains(FormUID) == true)
                    {
                        returnValue = true;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Check_ValidateForm_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }

            return returnValue;
        }

        /// <summary>
        /// AddOn 추가 폼 생성
        /// </summary>
        /// <param name="pVal"></param>
        /// <param name="pBaseClass"></param>
		private void Create_USERForm(SAPbouiCOM.MenuEvent pVal, ref PSH_BaseClass pBaseClass)
        {
            SAPbouiCOM.ProgressBar ProgBar01 = null;

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false); //MainMenu 클릭 후 화면이 열릴 때까지 Waiting 시작

                if (pVal.BeforeAction == true)
                {
                    for (int i = 0; i < PSH_Globals.classAllList.Count; i++)
                    {
                        if (PSH_Globals.classAllList[i].Name == pVal.MenuUID)
                        {
                            Type type = Type.GetType("PSH_BOne_AddOn." + pVal.MenuUID); //MenuUID와 동일한 클래스 Type 생성
                            dynamic baseClass = Activator.CreateInstance(type); //MenuUID와 동일한 클래스 Instance 생성
                            pBaseClass = baseClass; //PSH_BaseClass로 형변환
                            pBaseClass.LoadForm(""); //MenuUID에 클릭한 클래스의 LoadForm 호출
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Create_USERForm_Error: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop(); //MainMenu 클릭 후 화면이 열리면 Waiting 종료
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }

                if (pBaseClass.oForm != null)
                {
                    pBaseClass.oForm.Select(); //ProgressBar실행 종료후 현재 열린 화면으로 Focus 강제 이동
                }
            }
        }

        /// <summary>
        /// 시스템폼 생성
        /// </summary>
        /// <param name="pVal"></param>
        /// <param name="pBaseClass"></param>
        private void Create_SYSTEMForm(SAPbouiCOM.ItemEvent pVal, ref PSH_BaseClass pBaseClass)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    {
                        for (int i = 0; i < PSH_Globals.classAllList.Count; i++)
                        {
                           if (PSH_Globals.classAllList[i].Name == "S" + pVal.FormTypeEx) //접두어 "S" 포함
                            {
                                Type type = Type.GetType("PSH_BOne_AddOn.Core.S" + pVal.FormTypeEx); //Core폼과 동일한 클래스 Type 생성
                                dynamic baseClass = Activator.CreateInstance(type); //Core폼과 동일한 클래스 Instance 생성
                                pBaseClass = baseClass; //PSH_BaseClass로 형변환
                                pBaseClass.LoadForm(pVal.FormUID); //클래스의 LoadForm 호출
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("Create_SYSTEMForm_Error: " + ex.Message, BoMessageTime.bmt_Short, true);
            }
        }

        #region 이벤트

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
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                    Create_USERForm(pVal, ref oTempClass);
                    RecordSet01.DoQuery("EXEC Z_PS_FormCount '" + dataHelpClass.User_MSTCOD() + "','" + pVal.MenuUID + "'"); //Form 실행 횟수 저장

                    //C#Migration완료 후 제거 요망_S
                    RecordSet01.DoQuery("EXEC Z_PS_FormType '" + pVal.MenuUID + "'");

                    if (RecordSet01.Fields.Item("FormType").Value == "H")
                    {
                        RecordSet01.DoQuery("EXEC Z_PS_FormCount '" + dataHelpClass.User_MSTCOD() + "','" + pVal.MenuUID + "'"); //Form 실행 횟수 저장
                    }
                    //C#Migration완료 후 제거 요망_E
                }

                string FormUID = PSH_Globals.SBO_Application.Forms.ActiveForm.UniqueID;

                if (Check_ValidateForm(FormUID) == true)
                {
                    oTempClass = (PSH_BaseClass)PSH_Globals.ClassList[FormUID];
                    if (oTempClass.oForm == null)
                    {
                        return;
                    }
                    oTempClass.Raise_FormMenuEvent(FormUID, ref pVal, ref BubbleEvent);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("SBO_Application_MenuEvent_Error: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }
        }

        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            PSH_BaseClass oTempClass = new PSH_BaseClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                    Create_SYSTEMForm(pVal, ref oTempClass);
                }

                if (Check_ValidateForm(FormUID))
                {
                    oTempClass = (PSH_BaseClass)PSH_Globals.ClassList[FormUID];
                    if (oTempClass.oForm == null)
                    {
                        return;
                    }
                    oTempClass.Raise_FormItemEvent(FormUID, ref pVal, ref BubbleEvent);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("SBO_Application_ItemEvent_Error : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        private void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            PSH_BaseClass oTempClass;
            string FormUID;

            try
            {
                FormUID = BusinessObjectInfo.FormUID;

                if (Check_ValidateForm(FormUID))
                {
                    oTempClass = (PSH_BaseClass)PSH_Globals.ClassList[FormUID];
                    if (oTempClass.oForm == null)
                    {
                        return;
                    }
                    oTempClass.Raise_FormDataEvent(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("SBO_Application_FormDataEvent_Error : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        private void SBO_Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            PSH_BaseClass oTempClass;
            string FormUID;

            try
            {
                FormUID = eventInfo.FormUID;

                if (Check_ValidateForm(FormUID))
                {
                    oTempClass = (PSH_BaseClass)PSH_Globals.ClassList[FormUID];
                    if (oTempClass.oForm == null)
                    {
                        return;
                    }
                    oTempClass.Raise_RightClickEvent(FormUID, ref eventInfo, ref BubbleEvent);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("SBO_Application_RightClickEvent_Error : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        #endregion
    }
}




