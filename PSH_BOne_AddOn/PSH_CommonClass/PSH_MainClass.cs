using Microsoft.VisualBasic;
using System;
using System.Windows.Forms;
using SAPbouiCOM;
using Scripting;
using PSH_BOne_AddOn.Code;

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
        public void Initialize_ODBC_Variable()
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
                if (codeHelpClass.Right(ServerName, 3) == "223"){
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
                Query01 = "select UniqueID from [Authority_Screen] where Gubun ='H' and updateYN ='Y'and UserID ='" + PSH_Globals.oCompany.UserName + "'";
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
            string NowSeq;
            string AfSeq;
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
        /// 유효한 폼인지 검사
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
        /// AddOn 추가 폼 생성
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
                        #region 인사 관리
                        case "PH_PY001": //사원마스터등록

                            pBaseClass = new PH_PY001();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY005": //사업장정보등록  

                            pBaseClass = new PH_PY005();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY006": //승호작업등록

                            pBaseClass = new PH_PY006();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY011": //전문직호칭일괄변경

                            pBaseClass = new PH_PY011();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY017": //월근태집계

                            pBaseClass = new PH_PY017();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY021": //사원비상연락처관리

                            pBaseClass = new PH_PY021();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY201": //정년임박자 휴가경비 등록

                            pBaseClass = new PH_PY201();
                            pBaseClass.LoadForm("");
                            break;


                        case "PH_PY204": //교육계획등록

                            pBaseClass = new PH_PY204();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY203": //교육실적등록

                            pBaseClass = new PH_PY203();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY205": //교육계획VS실적조회

                            pBaseClass = new PH_PY205();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY009": //기찰기파일UPLOAD

                            pBaseClass = new PH_PY009();
                            pBaseClass.LoadForm("");
                            break;
						
						case "PH_PY202": //정년임박자 휴가경비 등록 현황

                            pBaseClass = new PH_PY202();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY510": //사원명부

                            pBaseClass = new PH_PY510();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY522": //임금피크 대상자 현황

                            pBaseClass = new PH_PY522();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY523": //임금피크 대상자월별 차수현황

                            pBaseClass = new PH_PY523();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY524": //퇴직금 중간정산내역

                            pBaseClass = new PH_PY524();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY419": //표준세액적용대상자등록

                            pBaseClass = new PH_PY419();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY016": //기본업무등록

                            pBaseClass = new PH_PY016();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY775": //개인별 연차현황

                            pBaseClass = new PH_PY775();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY776": //잔여연차현황(퇴직용)

                            pBaseClass = new PH_PY776();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA65": //연차현황(집계)

                            pBaseClass = new PH_PYA65();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY570": //연장/휴일근무현황

                            pBaseClass = new PH_PY570();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY585": //일일출근기록부

                            pBaseClass = new PH_PY585();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY610": //일일출근기록부

                            pBaseClass = new PH_PY610();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY615": //당직근무현황
                            pBaseClass = new PH_PY615();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY620": //연봉제휴일근무자현황

                            pBaseClass = new PH_PY620();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY675": //근무편성현황

                            pBaseClass = new PH_PY675();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA60": //학자금신청내역(집계)

                            pBaseClass = new PH_PYA60();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY625": //세탁자명부

                            pBaseClass = new PH_PY625();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY630": //세탁자명부

                            pBaseClass = new PH_PY630();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY700": //급여지급대장

                            pBaseClass = new PH_PY700();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY301": //학자금신청등록

                            pBaseClass = new PH_PY301();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY701": //급여지급대장(노조용)

                            pBaseClass = new PH_PY701();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA20": //급여부서별집계대장(부서)

                            pBaseClass = new PH_PYA20();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA40": //상여부서별집계대장(부서)

                            pBaseClass = new PH_PYA40();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA50": //DC전환자부담금지급내역

                            pBaseClass = new PH_PYA50();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA75": //교통비외수당지급대장

                            pBaseClass = new PH_PYA75();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY765": //급여증감내역서

                            pBaseClass = new PH_PY765();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY680": //상벌현황

                            pBaseClass = new PH_PY680();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY860": //호봉표조회

                            pBaseClass = new PH_PY860();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY508": //재직증명 등록 및 발급

                            pBaseClass = new PH_PY508();
                            pBaseClass.LoadForm("");
                            break;
                            
                        case "PH_PY770": //퇴직소득원천징수영수증출력

                            pBaseClass = new PH_PY770();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY780": //월고용보험내역

                            pBaseClass = new PH_PY780();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY785": //월국민연금내역

                            pBaseClass = new PH_PY785();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY790": //월건강보험내역

                            pBaseClass = new PH_PY790();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY795": //연간부서별급여현황  

                            pBaseClass = new PH_PY795();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY805": //급여수당변동내역

                            pBaseClass = new PH_PY805();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY810": //직급별통상임금내역

                            pBaseClass = new PH_PY810();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY815": //전체평균임금내역 

                            pBaseClass = new PH_PY815();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY820": //통상임금내역

                            pBaseClass = new PH_PY820();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY825": // 전문직O/T현황

                            pBaseClass = new PH_PY825();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY830": // 부서별인건비현황(기획)

                            pBaseClass = new PH_PY830();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY835": // 직급별O/T및수당현황

                            pBaseClass = new PH_PY835();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY840": // 풍산전자공시자료

                            pBaseClass = new PH_PY840();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY845": // 기간별급여지급내역

                            pBaseClass = new PH_PY845();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY850": // 소급분지급명세서

                            pBaseClass = new PH_PY850();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY855": // 개인별임금지급대장

                            pBaseClass = new PH_PY855();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY865": // 고용보험현황(계산용)
                            pBaseClass = new PH_PY865();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY870": // 담당별월O/T및수당현황   
                            pBaseClass = new PH_PY870();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY875": // 직급별수당집계대장  
                            pBaseClass = new PH_PY875();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY716": // 기간별급여부서별집계대장
                            pBaseClass = new PH_PY716();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY721": // 기간별상여부서별집계대장
                            pBaseClass = new PH_PY721();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY717": // 급여반별집계대장(기획용)
                            pBaseClass = new PH_PY717();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY715": // 급여부서별집계대장
                            pBaseClass = new PH_PY715();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY720": // 상여부서별집계대장
                            pBaseClass = new PH_PY720();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY725": // 급여직급별집계대장 
                            pBaseClass = new PH_PY725();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY740": // 상여직급별집계대장
                            pBaseClass = new PH_PY740();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY745": // 연간지급대장   
                            pBaseClass = new PH_PY745();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY750": // 근로소득징수현황
                            pBaseClass = new PH_PY750();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY755": // 동호회가입현황
                            pBaseClass = new PH_PY755();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY760": // 평균임금및퇴직금산출내역서
                            pBaseClass = new PH_PY760();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY019": // 반변경등록

                            pBaseClass = new PH_PY019();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY018": // 연봉제휴일교통비체크

                            pBaseClass = new PH_PY018();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY117": // 급상여마감작업

                            pBaseClass = new PH_PY117();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY123": // 가압류등록

                            pBaseClass = new PH_PY123();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY409": // 기부금조정명세자료등록

                            pBaseClass = new PH_PY409();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY555": // 일일근무자현황

                            pBaseClass = new PH_PY555();
                            pBaseClass.LoadForm("");
                            break;

						case "PH_PY010": //일일근태처리

                            pBaseClass = new PH_PY010();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY013": //위해코드등록(기계)

                            pBaseClass = new PH_PY013();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY014": //위해일수수정

                            pBaseClass = new PH_PY014();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY583": //근태마감체크

                            pBaseClass = new PH_PY583();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY120": //소급분급여생성

                            pBaseClass = new PH_PY120();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY133": //연봉제 횟차 관리

                            pBaseClass = new PH_PY133();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY119": //급상여은행파일생성

                            pBaseClass = new PH_PY119();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY002": //근태시간구분등록

                            pBaseClass = new PH_PY002();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY101": //보험률 등록

                            pBaseClass = new PH_PY101();
                            pBaseClass.LoadForm("");
                            break;
                            
                         case "PH_PY134": //소득세 / 주민세 조정

                            pBaseClass = new PH_PY134();
                            pBaseClass.LoadForm("");
                            break;
                            
                        case "PH_PY100": //기준세액설정

                            pBaseClass = new PH_PY100();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY109_1": //급상여변동자료항목수정

                            pBaseClass = new PH_PY109_1();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY131": //급상여변동자료항목수정

                            pBaseClass = new PH_PY131();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY129": //개인별 퇴직연금(DC형) 계산

                            pBaseClass = new PH_PY129();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY314": //대부금계산 내역 조회(급여변동자료용)

                            pBaseClass = new PH_PY314();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY695": //인사기록카드

                            pBaseClass = new PH_PY695();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY605": //근속보전휴가발생및사용내역

                            pBaseClass = new PH_PY605();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY560": //일출근현황

                            pBaseClass = new PH_PY560();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY565": //연장근무자현황

                            pBaseClass = new PH_PY565();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY575": //근태기찰현황

                            pBaseClass = new PH_PY575();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY580": //개인별근태월보

                            pBaseClass = new PH_PY580();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY985": //대부금계산 내역 조회(급여변동자료용)

                            pBaseClass = new PH_PY985();
                            pBaseClass.LoadForm();
                            break;


                        case "PH_PY590": //기간별근태집계표

                            pBaseClass = new PH_PY590();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY600": //일자별연장근무현황

                            pBaseClass = new PH_PY600();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY705": //교통비지급근태확인
                            pBaseClass = new PH_PY705();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY676": //근태시간내역조회
                            pBaseClass = new PH_PY676();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY679": //개인별근태자료집계
                            pBaseClass = new PH_PY679();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY681": //비근무일수현황
                            pBaseClass = new PH_PY681();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY645": //자격수당지급현황
                            pBaseClass = new PH_PY645();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA55": //정산징수및환급대장(집계)
                            pBaseClass = new PH_PYA55();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY550": //전체인원현황
                            pBaseClass = new PH_PY550();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY650": //노동조합간부현황
                            pBaseClass = new PH_PY650();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY685": //포상가급현황
                            pBaseClass = new PH_PY685();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY690": //생일자현황
                            pBaseClass = new PH_PY690();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA70": //소득세원천징수세액조정신청서출력
                            pBaseClass = new PH_PYA70();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY501": //여권발급현황
                            pBaseClass = new PH_PY501();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY525": //학력별인원현황
                            pBaseClass = new PH_PY525();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY935": //정기승호현황
                            pBaseClass = new PH_PY935();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY530": //연령별인원현황
                            pBaseClass = new PH_PY530();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY505": //입사자대장
                            pBaseClass = new PH_PY505();
                            pBaseClass.LoadForm(""); 
                            break;

                        case "PH_PY520": //퇴직및퇴직예정자대장
                            pBaseClass = new PH_PY520(); 
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY640": //국민연금퇴직전환금현황
                            pBaseClass = new PH_PY640();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY545": //인원현황(대내용)
                            pBaseClass = new PH_PY545();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY655": //보훈대상자현황
                            pBaseClass = new PH_PY655();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY660": //장애근로자현황
                            pBaseClass = new PH_PY660(); 
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY540": //인원현황(대외용)
                            pBaseClass = new PH_PY540();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY551": //평균인원조회
                            pBaseClass = new PH_PY551();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY535": //근속년수별인원현황
                            pBaseClass = new PH_PY535();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY665": //사원자녀현황
                            pBaseClass = new PH_PY665();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY670": //개인별차량현황
                            pBaseClass = new PH_PY670();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY503": //승진대상자명부
                            pBaseClass = new PH_PY503();
                            pBaseClass.LoadForm(""); 
                            break;

                        case "PH_PY507": //휴직자현황
                            pBaseClass = new PH_PY507(); 
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY635": //여행,교육자현황
                            pBaseClass = new PH_PY635();
                            pBaseClass.LoadForm(""); 
                            break;

                        case "PH_PY515": //재직자사원명부
                            pBaseClass = new PH_PY515();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY003": //근태월력등록

                            pBaseClass = new PH_PY003();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY595": //근속년수현황
                            pBaseClass = new PH_PY595();
                            pBaseClass.LoadForm(""); 
                            break;

                        case "PH_PY931": //표준세액적용대상자조회
                            pBaseClass = new PH_PY931();
                            pBaseClass.LoadForm(""); 
                            break;

                        case "PH_PY932": //전근무지등록현황
                            pBaseClass = new PH_PY932();
                            pBaseClass.LoadForm(""); 
                            break;

                        case "PH_PY933": //보수총액신고기초자료
                            pBaseClass = new PH_PY933();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY800": //인건비지급자료
                            pBaseClass = new PH_PY800();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY004": //근무조편성등록

                            pBaseClass = new PH_PY004();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA30": //상여지급대장(부서)
                            pBaseClass = new PH_PYA30();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA10": //급여지급대장(부서)
                            pBaseClass = new PH_PYA10();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY915": //근로소득원천징수부출력
                            pBaseClass = new PH_PY915();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY735": //상여봉투출력
                            pBaseClass = new PH_PY735();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY730": //급여봉투출력
                            pBaseClass = new PH_PY730();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY710": //상여지급대장
                            pBaseClass = new PH_PY710();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY930": //정산징수및환급대장
                            pBaseClass = new PH_PY930();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY925": //기부금명세서출력
                            pBaseClass = new PH_PY925();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY718": //생산완료금액대비O/T현황
                            pBaseClass = new PH_PY718();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY920": //원천징수영수증출력
                            pBaseClass = new PH_PY920();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY311": //통근버스운행등록

                            pBaseClass = new PH_PY311();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY910": //소득공제신고서출력
                            pBaseClass = new PH_PY910();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY400": //연멀정산대상자등록
                            pBaseClass = new PH_PY400();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY401": //전근무지등록
                            pBaseClass = new PH_PY401();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY402": //정산기초등록
                            pBaseClass = new PH_PY402();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY405": //의료비자료등록
                            pBaseClass = new PH_PY405();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY407": //정산기부금등록
                            pBaseClass = new PH_PY407();
                            pBaseClass.LoadForm(""); 
                            break;

                        case "PH_PY411": //연금저축등소득공제등록
                            pBaseClass = new PH_PY411();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY413": //월세액.주택임차차입금자료 등록
                            pBaseClass = new PH_PY413();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY113": //급(상)여 분개장 생성
                            pBaseClass = new PH_PY113();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY980": //근로소득지급명세서자료 전산매체수록
                            pBaseClass = new PH_PY980();
                            pBaseClass.LoadForm();
                            break;

                        case "PH_PY995": //퇴직소득지급명세서자료 전산매체수록
                            pBaseClass = new PH_PY995();
                            pBaseClass.LoadForm();
                            break;

                        case "PH_PY677": //일일근태이상자조회
                            pBaseClass = new PH_PY677();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY678": //당직근무자일괄등록
                            pBaseClass = new PH_PY678();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY683": //교대근무인정현황
                            pBaseClass = new PH_PY683();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY307": //학자금신청내역(분기별)
                            pBaseClass = new PH_PY307();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY306": //학자금신청내역(개인별)
                            pBaseClass = new PH_PY306();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY305": //학자금신청서
                            pBaseClass = new PH_PY305();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY315": //개인별 대부금 잔액현황
                            pBaseClass = new PH_PY315();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY313": //대부금계산
                            pBaseClass = new PH_PY313();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY107": //급상여기준일설정
                            pBaseClass = new PH_PY107();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY122": //급상여출력_개인부서설정등록
                            pBaseClass = new PH_PY122();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY015": //연차이월등록
                            pBaseClass = new PH_PY015();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY111": //급상여계산
                            pBaseClass = new PH_PY111();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY103": //공제항목설정
                            pBaseClass = new PH_PY103();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY312": //개인별버스요금등록(창원)
                            pBaseClass = new PH_PY312();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY121": //개인별 평가가급액 등록
                            pBaseClass = new PH_PY121();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY415": //소득정산계산
                            pBaseClass = new PH_PY415();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY302": //학자금지급완료처리
                            pBaseClass = new PH_PY302();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY303": //학자금은행파일생성
                            pBaseClass = new PH_PY303();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY417": //학자금은행파일생성
                            pBaseClass = new PH_PY417();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY007": //유류단가등록
                            pBaseClass = new PH_PY007();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY109": //급상여변동자료등록
                            pBaseClass = new PH_PY109();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY000": //사원마스터등록
                            pBaseClass = new PH_PY000();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY008": //일근태등록
                            pBaseClass = new PH_PY008();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY030": //공용등록
                            pBaseClass = new PH_PY030();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY031": //출장등록
                            pBaseClass = new PH_PY031();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY020": // 근태기본업무 변경등록(N.G.Y)_기계사업부
                            pBaseClass = new PH_PY020();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY104": // 고정수당공제금액일괄등록
                            pBaseClass = new PH_PY104();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY132": // 성과급 차등 개인별 계산
                            pBaseClass = new PH_PY132();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY116": // 퇴직금분개생
                            pBaseClass = new PH_PY116();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY105": // 호봉표등록
                            pBaseClass = new PH_PY105();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY108": //상여지급률설정
                            pBaseClass = new PH_PY108();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY125": //개인별 퇴직연금 설정등록(엑셀 Upload)
                            pBaseClass = new PH_PY125();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY102": //수당항목설정
                            pBaseClass = new PH_PY102();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY130": //팀별 성과급차등 등급등록
                            pBaseClass = new PH_PY130();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY032": //사용외출등록
                            pBaseClass = new PH_PY032();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY110": //개인별상여율등록
                            pBaseClass = new PH_PY110();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY106": //수당계산식설정
                            pBaseClass = new PH_PY106();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY114": //수당계산식설정
                            pBaseClass = new PH_PY114();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY112": //수당계산식설정
                            pBaseClass = new PH_PY112();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY309": //대부금등록
                            pBaseClass = new PH_PY309();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY034": //공용분개처리
                            pBaseClass = new PH_PY034();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY124": //베네피아 금액등록
                            pBaseClass = new PH_PY124();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY999": //메뉴권한관리
                            pBaseClass = new PH_PY999();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA80": //근무시간표출력
                            pBaseClass = new PH_PYA80();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PYA90": //근로소득간이지급명세서(세무서신고파일생성)
                            pBaseClass = new PH_PYA90();
                            pBaseClass.LoadForm();
                            break;

                        case "PH_PY526": //임금피크인원현황
                            pBaseClass = new PH_PY526();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY127": //개인별 4대보험 보수월액 및 정산금액 등록(엑셀 Upload)
                            pBaseClass = new PH_PY127();
                            pBaseClass.LoadForm("");
                            break;
                            
                        case "PH_PY310": //대부금개별상환(2019.11.21 송명규)
    						pBaseClass = new PH_PY310();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY115": //퇴직금계산(2019.11.22 송명규)
							pBaseClass = new PH_PY115();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY118": //급상여 E-Mail 발송(2019.12.16 송명규)
							pBaseClass = new PH_PY118();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY135": //급상여분개처리(2019.12.30 송명규)
                            pBaseClass = new PH_PY135();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY136": //급상여분개처리 배부규칙설정(2020.02.06 송명규)
                            pBaseClass = new PH_PY136();
                            pBaseClass.LoadForm("");
                            break;
                        #endregion 관리

                        #region 운영 관리
                        case "PS_DateChange": //날짜변경처리
                            pBaseClass = new PS_DateChange();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_DateCommit": //날짜변경승인
                            pBaseClass = new PS_DateCommit();
                            pBaseClass.LoadForm("");
                            break;

                        case "PH_PY998": //사용자 권한 조회
                            pBaseClass = new PH_PY998();
                            pBaseClass.LoadForm("");
                            break;
                        #endregion

                        #region 재무 관리
                        case "PS_CO685": //결산분개 등록
                            pBaseClass = new PS_CO685();
                            pBaseClass.LoadForm();
                            break;

                        case "PS_CO003": //결산분개 계정관리
                            pBaseClass = new PS_CO003();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI180": // 보조원장
                            pBaseClass = new PS_FI180();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI190": // 보조원장조회(EXCEL다운)
                            pBaseClass = new PS_FI190();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI110": // 보조재료비현황
                            pBaseClass = new PS_FI110();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI120": // 노무비현황
                            pBaseClass = new PS_FI120();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI130": // 제조경비현황
                            pBaseClass = new PS_FI130();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI140": // 판관비현황
                            pBaseClass = new PS_FI140();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI150": // 영업외손익.특별손익명세서
                            pBaseClass = new PS_FI150();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI160": // 외상매입/미지급금 잔액현황
                            pBaseClass = new PS_FI160();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI170": // 신용카드 사용내역
                            pBaseClass = new PS_FI170();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI913": // 사원별 접대비 명세서
                            pBaseClass = new PS_FI913();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI924": // 거래처별 회전일 현황
                            pBaseClass = new PS_FI924();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI909": // 미수현황
                            pBaseClass = new PS_FI909();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI910": // 매출채권 연령분석(받을어음)
                            pBaseClass = new PS_FI910();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI931": // 부가세집계표
                            pBaseClass = new PS_FI931();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI420": // 분개전표 연결발행
                            pBaseClass = new PS_FI420();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI220": // 합계잔액시산표
                            pBaseClass = new PS_FI220();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI904": // 고정자산현황-IFRS
                            pBaseClass = new PS_FI904();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI918": // 대체출고대장
                            pBaseClass = new PS_FI918();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI912": // 예산집행현황
                            pBaseClass = new PS_FI912();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI210": //원가요소(배부담당)별비용집계대장
                            pBaseClass = new PS_FI210();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI215": //지급현황
                            pBaseClass = new PS_FI215();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI960": //입,퇴사자현황
                            pBaseClass = new PS_FI960();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI961": //월별계정별비용현황
                            pBaseClass = new PS_FI961();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI900": //IFRS - 시산표 추출
                            pBaseClass = new PS_FI900();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI901": //IFRS - 매입매출채무채권추출
                            pBaseClass = new PS_FI901();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI902": //IFRS - NRV 평가
                            pBaseClass = new PS_FI902();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI929": //IFRS - 재무보고서
                            pBaseClass = new PS_FI929();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI932": //차입금등록
                            pBaseClass = new PS_FI932();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI933": //차입금대장
                            pBaseClass = new PS_FI933();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI421": //입금표등록
                            pBaseClass = new PS_FI421();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO090": //통계주요지표 값 입력
                            pBaseClass = new PS_CO090();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO080": //코스트센터비용집계
                            pBaseClass = new PS_CO080();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO070": //활동량계산
                            pBaseClass = new PS_CO070();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO185": //담당별 판매/ 원재료 / rm단가등록
                            pBaseClass = new PS_CO185();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO170": //재공완료품대체관리등록
                            pBaseClass = new PS_CO170();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI100": //계정별보조원장
                            pBaseClass = new PS_FI100();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI090": //월계표
                            pBaseClass = new PS_FI090();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO110": //배부사이클실행
                            pBaseClass = new PS_CO110();
                            pBaseClass.LoadForm();
                            break;

                        case "PS_CO260": //연차설정등록
                            pBaseClass = new PS_CO260();
                            pBaseClass.LoadForm();
                            break;

                        case "PS_CO250": //개인별 퇴충계산
                            pBaseClass = new PS_CO250();
                            pBaseClass.LoadForm();
                            break;

                        case "PS_CO160": //재공 원가 이동등록
                            pBaseClass = new PS_CO160();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO050": //활동유형등록
                            pBaseClass = new PS_CO050();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO060": //통계주요지표등록
                            pBaseClass = new PS_CO060();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO040": //코스트센터그룹등록
                            pBaseClass = new PS_CO040();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO100": //배부사이클생성
                            pBaseClass = new PS_CO100();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO020": //원가요소그룹등록
                            pBaseClass = new PS_CO020();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO010": //원가요소등록
                            pBaseClass = new PS_CO010();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO001": //결산마감관리
                            pBaseClass = new PS_CO001();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO658": //통합재무제표 계정 관리
                            pBaseClass = new PS_CO658();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO600": //통합재무제표
                            pBaseClass = new PS_CO600();
							pBaseClass.LoadForm("");
                            break;

                        case "PS_CO605": //통합재무제표
                            pBaseClass = new PS_CO605();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO606": //통합재무제표 본사전송
                            pBaseClass = new PS_CO606();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO610": //고정자산 본계정 대체
                            pBaseClass = new PS_CO610();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO660": //기간비용등록
                            pBaseClass = new PS_CO660();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO665": //기간비용현황(연간)
                            pBaseClass = new PS_CO665();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO670": //기간비용분개등록
                            pBaseClass = new PS_CO670();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO120": //공정별 원가계산
                            pBaseClass = new PS_CO120();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO130": //제품별 원가계산
                            pBaseClass = new PS_CO130();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO131": //원가계산재공현황
                            pBaseClass = new PS_CO131();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO510": //원가계산사전점검조회
                            pBaseClass = new PS_CO510();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO520": //제품생산 원가항목별 조회
                            pBaseClass = new PS_CO520();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO501": //품목별원가등록
                            pBaseClass = new PS_CO501();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO502": //품목별평균원가항목등록
                            pBaseClass = new PS_CO502();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO503": //일일가득액및생산원가계산
                            pBaseClass = new PS_CO503();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO504": //일일판매및생산집계
                            pBaseClass = new PS_CO504();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO210": //5.휘팅제품원가계산
                            pBaseClass = new PS_CO210();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO901": //월별 이익률 분석(기계)
                            pBaseClass = new PS_CO901();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO920": //월별 판매실적 대비 공수 조회(기계)
                            pBaseClass = new PS_CO920();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO921": //월별 작번별 수주금액 대비 공수 현황(기계)
                            pBaseClass = new PS_CO921();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO530": //분말부자재비용분석
                            pBaseClass = new PS_CO530();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO540": //원가 분석 레포트(안강)
                            pBaseClass = new PS_CO540();
                            pBaseClass.LoadForm("") ;
                            break;

                        case "PS_MM921": //부자재불출대장(부서)
                            pBaseClass = new PS_MM921();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FX241": //고정자산현황(부서)
                            pBaseClass = new PS_FX241();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI030": //어음발행리스트
                            pBaseClass = new PS_FI030();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI040": // 만기일자별 받을어음현황
                             pBaseClass = new PS_FI040();
                             pBaseClass.LoadForm("");
                             break;

                        case "PS_FI050": //받을어음 추심의뢰서
                            pBaseClass = new PS_FI050();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI060": //받을어음 수탁(발송)대장
                            pBaseClass = new PS_FI060();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI070": //기간별 지급어음현황
                            pBaseClass = new PS_FI070();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI907": //부도어음현황
                            pBaseClass = new PS_FI907();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI080": //일계표
                            pBaseClass = new PS_FI080();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_FI963": //보조원장 거래처명 변경화
                            pBaseClass = new PS_FI963();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_CO800": //제품원재료 변환
                            pBaseClass = new PS_CO800();
                            pBaseClass.LoadForm("");
                            break;
                        #endregion

                        #region 생산관리
                        case "PS_PP048": //스크랩입고등록
                            pBaseClass = new PS_PP048();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_PP047": //금속분말 재작업등록
                            pBaseClass = new PS_PP047();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_PP730": //분말일자별작업조회
                            pBaseClass = new PS_PP730();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_PP084": //생산완료등록(분말)
                            pBaseClass = new PS_PP084();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_PP036": //MG생산일보조회
                            pBaseClass = new PS_PP036();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_PP097": //분말검사결과확인등록
                            pBaseClass = new PS_PP097();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_PP040": //작업일보등록(작지)
                            pBaseClass = new PS_PP040();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_PP030": //작업지시등록
                            pBaseClass = new PS_PP030();
                            pBaseClass.LoadForm("");
                            break;

                        #endregion

                        #region 판매관리
                        case "PS_SD600": //견적양식등록
                            pBaseClass = new PS_SD600();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD602": //견적등록
                            pBaseClass = new PS_SD602();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD603": //견적조회
                            pBaseClass = new PS_SD603();
                            pBaseClass.LoadForm();
                            break;

                        case "PS_SD901": //(기계)판매오더 수주처 변경시 작업시시 변경처리
                            pBaseClass = new PS_SD901();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD012": //파카 라벨 출력
                            pBaseClass = new PS_SD012();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD013": //AR송장 만기일 조회 및 승인
                            pBaseClass = new PS_SD013();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD380": //납품명세서
                            pBaseClass = new PS_SD380();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD110": //견적및수주현황
                            pBaseClass = new PS_SD110();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD120": //견적서현황
                            pBaseClass = new PS_SD120();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD130": //수주접수현황
                            pBaseClass = new PS_SD130();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD160": //수주/견적비교현황
                            pBaseClass = new PS_SD160();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD180": //수주실적(거래처별)
                            pBaseClass = new PS_SD180();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD240": //미납품현황(품목별)
                            pBaseClass = new PS_SD240();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD241": //미납품현황(포장사업팀)
                            pBaseClass = new PS_SD241();
                            pBaseClass.LoadForm("");
                            break;
                            
                        case "PS_SD250": //제품재고현황
                            pBaseClass = new PS_SD250();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD270": //판매미확정현황
                            pBaseClass = new PS_SD270();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD280": //제품출고조회
                            pBaseClass = new PS_SD280();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD290": //납품지연현황
                            pBaseClass = new PS_SD290();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD295": //거래처별납기분석표
                            pBaseClass = new PS_SD295();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD310": //판매실적(납품처별)
                            pBaseClass = new PS_SD310();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD400": //판매현황(멀티)
                            pBaseClass = new PS_SD400();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD320": //품목별판매내역서
                            pBaseClass = new PS_SD320();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD330": //품목분류별월별판매현황
                            pBaseClass = new PS_SD330();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD471": //판매현황(거래처별)
                            pBaseClass = new PS_SD471();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD476": //판매현황(제품,규격별)
                            pBaseClass = new PS_SD476();
                            pBaseClass.LoadForm("");
                            break;
                            
                        case "PS_PP688": //생산의뢰서
                            pBaseClass = new PS_PP688();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_MM091": //기타출고현황
                            pBaseClass = new PS_MM091();
                            pBaseClass.LoadForm("");
                            break;
                            
                        case "PS_PP911": //작번이력현황
                            pBaseClass = new PS_PP911();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD510": //기성매출현황
                            pBaseClass = new PS_SD510();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD520": //기성매출현황
                            pBaseClass = new PS_SD520();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD190": //판매실적조회
                            pBaseClass = new PS_SD190();
                            pBaseClass.LoadForm("");
                            break;
                            
                        case "PS_SD950": //채권만기일관리
                            pBaseClass = new PS_SD950();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_PP325": //부품주문번호별출고대장
                            pBaseClass = new PS_PP325();
                            pBaseClass.LoadForm("");
                            break;

                        case "PS_SD285": //타계정출고현황(샘플,감모손)
                            pBaseClass = new PS_SD285();
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
                        {

                            //case "60100":       //인사관리>사원마스터데이터 (사용자 정의 필드)
                            //    pBaseClass = new SM60100();
                            //    pBaseClass.LoadForm(pVal.FormUID);
                            //    break;
                        }
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
        /// 생성된 폼 클래스 해제(사용안함:exe 메모리 해제 안됨, 2018.12.03 송명규)
        /// </summary>
        private void TerminateApplication()
        {
            int i;
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
            string FormUID;

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
                Create_SYSTEMForm(pVal, ref oTempClass);

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
                            else if (pVal.Before_Action == false) //FORM_UNLOAD 이벤트가 Before_Action == false 일 때는 PSH_Globals.ClassList[FormUID] 에 index 오류 발생하므로 강제 return
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
            PSH_BaseClass oTempClass;
            string FormUID;

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
            PSH_BaseClass oTempClass;
            string FormUID;

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




