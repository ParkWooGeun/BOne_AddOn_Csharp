using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using SAP.Middleware.Connector;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 구매견적
    /// </summary>
    internal class PS_MM010 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_MM010H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_MM010L; //등록라인
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private string oPurchase;
        private string oPQType;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM010.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_MM010_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_MM010");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.DataBrowser.BrowseBy = "DocNum";

                oForm.Freeze(true);
                if (!string.IsNullOrEmpty(oFormDocEntry))
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                }
                PS_MM010_CreateItems();
                PS_MM010_ComboBox_Setting();
                PS_MM010_FormClear();
                PS_MM010_Add_MatrixRow(0, true);

                oForm.EnableMenu("1283", false); // 삭제
                oForm.EnableMenu("1286", true); // 닫기
                oForm.EnableMenu("1287", false); // 복제
                oForm.EnableMenu("1284", true); // 취소
                oForm.EnableMenu("1293", true); // 행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                if (!string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_MM010_FormItemEnabled();
                    oForm.Items.Item("DocNum").Specific.Value = oFormDocEntry;
                    oForm.Items.Item("CntcCode").Specific.Value = "";
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else
                {
                    PS_MM010_Initialization();
                }
                oForm.Visible = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_MM010_CreateItems()
        {
            try
            {
                oDS_PS_MM010H = oForm.DataSources.DBDataSources.Item("@PS_MM010H");
                oDS_PS_MM010L = oForm.DataSources.DBDataSources.Item("@PS_MM010L");

                // 메트릭스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;

                oDS_PS_MM010H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_MM010_ComboBox_Setting()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.DataSources.UserDataSources.Add("SumQty", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("SumQty").Specific.DataBind.SetBound(true, "", "SumQty");

                oForm.DataSources.UserDataSources.Add("SumWeight", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("SumWeight").Specific.DataBind.SetBound(true, "", "SumWeight");

                // 사업장
                sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                //견적형태
                sQry = "SELECT Code, Name From [@PSH_RETYPE]";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("PQType").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oDS_PS_MM010H.SetValue("U_PQType", 0, "");

                oForm.Items.Item("RotateYN").Specific.ValidValues.Add("N", "[N]일반품");
                oForm.Items.Item("RotateYN").Specific.ValidValues.Add("Y", "[Y]순환품");
                oDS_PS_MM010H.SetValue("U_RotateYN", 0, "N");

                //구매방식, 품목구분
                sQry = "SELECT Code, Name From [@PSH_ORDTYP] Order by Code";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("Purchase").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oMat01.Columns.Item("ItemGpCd").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oDS_PS_MM010H.SetValue("U_Purchase", 0, "");

                //품목대분류
                sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Order by Code";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oMat01.Columns.Item("ItmBSort").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                //매입기준단위
                sQry = "SELECT Code, Name From [@PSH_UOMORG] Order by Code";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oMat01.Columns.Item("OBasUnit").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                //외주사유코드
                sQry = "  SELECT    T1.U_Minor AS [Code],";
                sQry += "           T1.U_CdName AS [Value]";
                sQry += " FROM      [@PS_SY001H] AS T0";
                sQry += "           INNER JOIN";
                sQry += "           [@PS_SY001L] AS T1";
                sQry += "               ON T0.Code = T1.Code";
                sQry += " WHERE     T0.Code = 'P201'";
                sQry += "           AND T1.U_UseYN = 'Y'";
                sQry += " ORDER BY  T1.U_Seq";

                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oMat01.Columns.Item("OutCode").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// PS_MM010_Initialization
        /// </summary>
        private void PS_MM010_Initialization()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                //아이디별 사번 세팅
                oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();

                //품목구분 종전입력값
                if (string.IsNullOrEmpty(oForm.Items.Item("Purchase").Specific.Value.ToString().Trim()) && !string.IsNullOrEmpty(oPurchase))
                {
                    oForm.Items.Item("Purchase").Specific.Select(oPurchase, SAPbouiCOM.BoSearchKey.psk_ByValue);
                }

                //견적형태 종전입력값
                if (string.IsNullOrEmpty(oForm.Items.Item("PQType").Specific.Value.ToString().Trim()) && !string.IsNullOrEmpty(oPQType))
                {
                    oForm.Items.Item("PQType").Specific.Select(oPQType, SAPbouiCOM.BoSearchKey.psk_ByValue);
                }

                oForm.Items.Item("RotateYN").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// R3 Interface 및 FTP를 통한 자료 업로드
        /// </summary>
        /// <param name="BPLId"></param>
        /// <param name="ItmBsort"></param>
        /// <param name="ItemCode"></param>
        /// <param name="ItemName"></param>
        /// <param name="Size"></param>
        /// <param name="Qty"></param>
        /// <param name="Unit"></param>
        /// <param name="RequestDate"></param>
        /// <param name="DueDate"></param>
        /// <param name="ItemType"></param>
        /// <param name="RequestNo"></param>
        /// <param name="BEDNR"></param>
        /// <param name="i"></param>
        /// <param name="LastRow"></param>
        /// <param name="FileName"></param>
        /// <param name="dir"></param>
        /// <returns></returns>
        private string PS_RFC_Sender(string BPLId, string ItmBsort, string ItemCode, string ItemName, string Size, double Qty, string Unit, string RequestDate, string DueDate, string ItemType, string RequestNo, string BEDNR, int i, int LastRow, string FileName, string dir)
        {
            string returnValue = string.Empty;
            string WERKS = string.Empty;
            string Rotate;
            string Comments;
            string errMessage = string.Empty;
            string Client; //클라이언트(운영용:210, 테스트용:710)
            string ServerIP;
            string errCode = string.Empty;
            RfcDestination rfcDest = null;
            RfcRepository rfcRep = null;
            IRfcFunction oFunction = null;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                //서버IP(운영용:192.1.11.3, 테스트용:192.1.11.7)
                //Real
                Client = "210";
                ServerIP = "192.1.11.3";

                ////test
                //Client = "810";
                //ServerIP = "192.1.11.7";

                //0. 연결
                if (dataHelpClass.SAPConnection(ref rfcDest, ref rfcRep, "PSC", ServerIP, Client, "ifuser", "pdauser") == false)
                {
                    errMessage = "풍산 SAP R3에 로그온 할 수 없습니다. 관리자에게 문의 하세요.";
                    throw new Exception();
                }

                if (oForm.Items.Item("RotateYN").Specific.Value == "Y")
                {
                    Rotate = "CB"; //순환품구매요청
                }
                else
                {
                    Rotate = "NB"; //표준구매요청
                }

                if ((ItmBsort == "20" && ItemCode.Length == 11) || ItmBsort == "50")
                {
                    oFunction = rfcRep.CreateFunction("ZMM_SUB_PR"); //부자재이면서 11자리 인경우 통구실 구매요청
                }
                else
                {
                    oFunction = rfcRep.CreateFunction("ZMM_INTF_GROUP"); //그외는 자재코드 맵핑해서 처리
                }

                switch (BPLId)
                {
                    case "1":
                        WERKS = "9200";
                        break;
                    case "2":
                        WERKS = "9300";
                        break;
                    case "3":
                        WERKS = "9500";
                        break;
                    case "5":
                        WERKS = "9600";
                        break;
                }

                Comments = oDS_PS_MM010L.GetValue("U_Comments", i).ToString().Trim(); //비고
                oFunction.SetValue("I_WERKS", WERKS); //플랜트 홀딩스 창원 9200, 홀딩스 부산 9300, 포장사업팀 9500, 포장온산 9600
                oFunction.SetValue("I_MATNR", ItemCode); //자재코드 char(18) 

                //통합구매로 본사R3에 값을 전달할 때 품목길이 Check (2012.03.28 송명규)
                string[] arrItemName = null;
                //ItemName의 길이가 40이하일 경우에만 I_MAKTX 에 전달
                if (ItemName.Length <= 40)
                {
                    oFunction.SetValue("I_MAKTX", ItemName); //자재내역 char(40)
                    oFunction.SetValue("I_WRKST", Size); //자재규격 char(48)
                }
                else //40 초과일 경우는 I_WRKST에 ItemName에서 ":" 뒤의 규격만 전달
                {
                    arrItemName = ItemName.Split(new char[] { ':' });

                    oFunction.SetValue("I_MAKTX", arrItemName[0]);
                    if (string.IsNullOrEmpty(Size))  //Size에 값이 없을 때만
                    {
                        oFunction.SetValue("I_WRKST", arrItemName[1]); //자재규격 char(48)
                    }
                    else  //Size에 값이 있을 때는 Size의 값을 전달
                    {
                        oFunction.SetValue("I_WRKST", Size); //자재규격 char(48)
                    }
                }
                //통합구매로 본사R3에 값을 전달할 때 품목길이 Check (2012.03.28 송명규)
                oFunction.SetValue("I_ZUSE", Comments); //비고 char(132)"
                oFunction.SetValue("I_MENGE", Qty); //구매요청수량 dec(13,3)
                oFunction.SetValue("I_MEINS", Unit); //단위 char(3)
                oFunction.SetValue("I_BADAT", RequestDate); //구매요청일 char(8)
                oFunction.SetValue("I_LFDAT", DueDate); //납품일 char(8)
                oFunction.SetValue("I_MATKL", ItemType); //자재그룹 char(9)
                oFunction.SetValue("I_ZBANFN", RequestNo); //구매요청번호
                oFunction.SetValue("I_BEDNR", BEDNR); //구매담당자(전화번호)
                oFunction.SetValue("I_BSART", Rotate); //순환품 구매요청

                errCode = "1"; // 아래 invoke 오류 체크를 위한 변수대입
                oFunction.Invoke(rfcDest); //Function 실행
                errCode = string.Empty;// 이상 없을 경우 초기화
                if (string.IsNullOrEmpty(oFunction.GetValue("E_MESSAGE").ToString()))
                {
                    returnValue = oFunction.GetValue("E_BANFN").ToString() + "/" + oFunction.GetValue("E_BNFPO").ToString(); //통합구매요청번호 '//통합구매요청 품목번호
                }
                else
                {
                    oDS_PS_MM010L.SetValue("U_MESSAGE", i, "");
                    oDS_PS_MM010L.SetValue("U_MESSAGE", i, oFunction.GetValue("E_MESSAGE").ToString());
                    errMessage = oFunction.GetValue("E_MESSAGE").ToString();
                    throw new Exception();
                }

                if (!string.IsNullOrEmpty(FileName))
                {
                    string ip = "192.1.11.3";
                    string port = "21";
                    string userid = "ftpadm";
                    string pwd = "psc1004";
                    string upLoadFile = dir;
                    if (dataHelpClass.FTPConn_Upload(ip, port, userid, pwd, "trans/group/", upLoadFile) == false)
                    {
                        errMessage = "FTP 업로드 오류 관리자에게 문의하세요.";
                        throw new Exception();
                    }
                    else
                    {
                        oFunction = rfcRep.CreateFunction("ZMM_INTF_GROUP_FILE");
                        if (oFunction == null)
                        {
                            oDS_PS_MM010L.SetValue("U_MESSAGE1", i, ""); //초기화
                            oDS_PS_MM010L.SetValue("U_MESSAGE1", i, "함수(ZMM_INTF_GROUP_FILE) 생성오류");
                            PSH_Globals.SBO_Application.MessageBox("함수(ZMM_INTF_GROUP_FILE) 생성오류.");
                        }
                        else
                        {
                            //함수정상호출 하면
                            oFunction.SetValue("I_MATNR", ItemCode); //품목코드
                            oFunction.SetValue("I_WERKS", WERKS); //사업장
                            oFunction.SetValue("I_FILENAME", FileName);

                            errCode = "2"; //ZMM_INTF_GROUP_FILE 함수호출 체크를 위한 변수값 대입
                            oFunction.Invoke(rfcDest); //Function 실행
                            errCode = string.Empty; // 이상 없을 경우 초기화

                            if (string.IsNullOrEmpty(oFunction.GetValue("E_MESSAGE").ToString()))
                            {
                                //함수호출 에러
                                oDS_PS_MM010L.SetValue("U_MESSAGE1", i, "");
                                oDS_PS_MM010L.SetValue("U_MESSAGE1", i, "통합구매(R/3) 함수(ZMM_INTF_GROUP_FILE)호출중 오류발생");
                                PSH_Globals.SBO_Application.MessageBox("통합구매(R/3) 함수(ZMM_INTF_GROUP_FILE)호출중 오류발생");
                            }
                            else
                            {
                                oDS_PS_MM010L.SetValue("U_MESSAGE1", i, "");
                                oDS_PS_MM010L.SetValue("U_MESSAGE1", i, oFunction.GetValue("E_MESSAGE").ToString());
                                PSH_Globals.SBO_Application.MessageBox(oFunction.GetValue("E_MESSAGE").ToString());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (errCode == "E1")
                {
                    PSH_Globals.SBO_Application.MessageBox("안강(R/3)서버 함수호출중 오류발생");
                }
                else if (errCode == "E2")
                {
                    PSH_Globals.SBO_Application.MessageBox("통합구매(R/3) 함수(ZMM_INTF_GROUP_FILE)호출중 오류발생");
                }
                else if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                if (rfcDest != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rfcDest);
                }

                if (rfcRep != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rfcRep);
                }
                
                if (oFunction != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oFunction);
                }
            }
            return returnValue;
        }

        /// <summary>
        /// 견적의뢰서 출력
        /// </summary>
        [STAThread]
        private void PS_MM010_Print_Report01()
        {
            string WinTitle;
            string ReportName;
            string oText = string.Empty;
            string DocNum;
            string Purchase;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                DocNum = oDS_PS_MM010H.GetValue("DocNum", 0).ToString().Trim();
                Purchase = oDS_PS_MM010H.GetValue("U_Purchase", 0).ToString().Trim();

                WinTitle = "[PS_MM010] 견적의뢰서(결재)";
                ReportName = "PS_MM010_01.rpt";
                //프로시저 : PS_MM010_01

                if (Purchase == "10") //품의구분:원재료(10)
                {
                    oText = "작번 / 품명 / Main 품명";
                }
                else if (Purchase == "20") //품의구분:부재료(20)
                {
                    oText = "사용처";
                }

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();

                // Formula
                dataPackFormula.Add(new PSH_DataPackClass("@F01", oText));

                // Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@DocNum", DocNum));

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 견적의뢰서(구매요청서) 출력
        /// </summary>
        [STAThread]
        private void PS_MM010_Print_Report02()
        {
            string WinTitle;
            string ReportName;
            string DocNum;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                DocNum = oDS_PS_MM010H.GetValue("DocNum", 0).ToString().Trim();

                WinTitle = "[PS_MM010] 구매요청서(통합구매)";
                ReportName = "PS_MM010_02.rpt";
                //프로시저 : PS_MM010_03

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

                // Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@DocNum", DocNum));

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_MM010_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("PQType").Enabled = true;
                    oForm.Items.Item("Purchase").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("PQType").Enabled = true;
                    oForm.Items.Item("Purchase").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = false;

                    if (oDS_PS_MM010H.GetValue("Status", 0).ToString().Trim() == "O") //문서상태:열기
                    {
                        if (oDS_PS_MM010H.GetValue("U_PQType", 0).ToString().Trim() == "10")
                        {
                            oForm.Items.Item("BPLId").Enabled = false;
                            oForm.Items.Item("CntcCode").Enabled = true;
                            oForm.Items.Item("PQType").Enabled = true;
                            oForm.Items.Item("Purchase").Enabled = false;
                            oForm.Items.Item("DocDate").Enabled = true;
                            oForm.Items.Item("Mat01").Enabled = true;
                        }
                        else if (oDS_PS_MM010H.GetValue("U_PQType", 0).ToString().Trim() == "20")
                        {
                            if (oDS_PS_MM010H.GetValue("U_RFCAdms", 0).ToString().Trim() == "Y")
                            {
                                oForm.Items.Item("BPLId").Enabled = false;
                                oForm.Items.Item("CntcCode").Enabled = false;
                                oForm.Items.Item("PQType").Enabled = false;
                                oForm.Items.Item("Purchase").Enabled = false;
                                oForm.Items.Item("DocDate").Enabled = false;
                                oForm.Items.Item("Mat01").Enabled = false;
                            }
                            else
                            {
                                oForm.Items.Item("BPLId").Enabled = true;
                                oForm.Items.Item("CntcCode").Enabled = true;
                                oForm.Items.Item("PQType").Enabled = true;
                                oForm.Items.Item("Purchase").Enabled = true;
                                oForm.Items.Item("DocDate").Enabled = true;
                                oForm.Items.Item("Mat01").Enabled = true;
                            }
                        }
                    }
                    else  //문서상태:닫기(취소)
                    {
                        oForm.Items.Item("BPLId").Enabled = false;
                        oForm.Items.Item("CntcCode").Enabled = false;
                        oForm.Items.Item("PQType").Enabled = false;
                        oForm.Items.Item("Purchase").Enabled = false;
                        oForm.Items.Item("DocDate").Enabled = false;
                        oForm.Items.Item("Mat01").Enabled = false;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_MM010_Add_MatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_MM010L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_MM010L.Offset = oRow;
                oDS_PS_MM010L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// HeaderSpaceLineDel
        /// </summary>
        /// <returns></returns>
        private bool PS_MM010_HeaderSpaceLineDel()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_MM010H.GetValue("U_BPLId", 0)))
                {
                    errMessage = "사업장은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM010H.GetValue("U_CntcCode", 0)))
                {
                    errMessage = "담당자는 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM010H.GetValue("U_PQType", 0)))
                {
                    errMessage = "견적형태는 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM010H.GetValue("U_Purchase", 0)))
                {
                    errMessage = "구매방식은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM010H.GetValue("U_DocDate", 0)))
                {
                    errMessage = "전기일은 필수사항입니다. 확인하세요.";
                    throw new Exception();
                }
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            return returnValue;
        }

        /// <summary>
        /// MatrixSpaceLineDel
        /// </summary>
        /// <returns></returns>
        private bool PS_MM010_MatrixSpaceLineDel()
        {
            bool returnValue = false;
            int i;
            string errMessage = string.Empty;

            try
            {
                oMat01.FlushToDataSource();

                // 라인
                if (oMat01.VisualRowCount == 0)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }
                else if (oMat01.VisualRowCount == 1 && string.IsNullOrEmpty(oDS_PS_MM010L.GetValue("U_CGNo", 0)))
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                {
                    if (string.IsNullOrEmpty(oDS_PS_MM010L.GetValue("U_CGNo", i)))
                    {
                        errMessage = "청구번호는 필수사항입니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (Convert.ToDouble(oDS_PS_MM010L.GetValue("U_Weight", i)) == 0)
                    {
                        errMessage = "중량은 필수사항입니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (oForm.Items.Item("PQType").Specific.Value == "20" && oForm.Items.Item("Purchase").Specific.Value == "20" && oDS_PS_MM010L.GetValue("U_ItemCode", i).Length != 11)
                    {
                        errMessage = "통합구매요청은 통합코드로만 가능합니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (PS_MM010_CheckDate(oMat01.Columns.Item("CGNo").Cells.Item(i + 1).Specific.Value) == false)
                    {
                        errMessage = i + 1 + "행 [" + oMat01.Columns.Item("ItemCode").Cells.Item(i + 1).Specific.Value + "]의 구매견적일은 구매요청일과 같거나 늦어야합니다. 확인하십시오." + "해당 견적은 전체가 등록되지 않습니다.";
                        throw new Exception();
                    }
                }
                oMat01.LoadFromDataSource();
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            return returnValue;
        }

        /// <summary>
        /// Delete_EmptyRow
        /// </summary>
        private void PS_MM010_Delete_EmptyRow()
        {
            int i;
            string errMessage = string.Empty;

            try
            {
                oMat01.FlushToDataSource();

                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oDS_PS_MM010L.GetValue("U_CGNo", i).ToString().Trim()))
                    {
                        oDS_PS_MM010L.RemoveRecord(i); // Mat01에 마지막라인(빈라인) 삭제
                    }
                }
                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
        }

        /// <summary>
        /// CheckDate
        /// </summary>
        /// <returns></returns>
        private bool PS_MM010_CheckDate(string pBaseEntry)
        {
            bool returnValue = false;
            string Query01;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                Query01 = " EXEC PS_Z_CHECK_DATE '";
                Query01 += pBaseEntry + "','"; // BaseEntry
                Query01 += "" + "','";  //BaseLine(빈값)
                Query01 += "PS_MM010" + "','"; //DocType
                Query01 += oForm.Items.Item("DocDate").Specific.Value.ToString().Trim() + "'"; //CurDocDate

                oRecordSet01.DoQuery(Query01);

                if (oRecordSet01.Fields.Item("ReturnValue").Value == "False")
                {
                    returnValue = false;
                }
                else
                {
                    returnValue = true;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
            return returnValue;
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_MM010_FormClear()
        {
            string DocNum;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocNum = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM010'", "");
                if (Convert.ToDouble(DocNum) == 0)
                {
                    oForm.Items.Item("DocNum").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocNum").Specific.Value = DocNum;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_MM010_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            int i;
            int SumQty = 0;
            string sQry;
            double SumWeight = 0;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                switch (oUID)
                {
                    case "BPLId":
                        //사업장에 따른 선진행부서 설정

                        sQry = "  SELECT      U_Code,";
                        sQry += "             U_CodeNm";
                        sQry += " FROM        [@PS_HR200L]";
                        sQry += "             WHERE Code = '1'";
                        sQry += "             AND U_Char2 = '" + oForm.Items.Item("BPLId").Specific.Value + "'";
                        sQry += "             AND U_UseYN = 'Y'";
                        sQry += "             AND U_Code NOT IN ('1100','2100')";
                        sQry += " ORDER BY    U_Seq";

                        if (oForm.Items.Item("BefTeam").Specific.ValidValues.Count > 0)
                        {
                            for (i = oForm.Items.Item("BefTeam").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                            {
                                oForm.Items.Item("BefTeam").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                        }

                        oForm.Items.Item("BefTeam").Specific.ValidValues.Add("", "");
                        dataHelpClass.Set_ComboList(oForm.Items.Item("BefTeam").Specific, sQry, "", false, false);

                        oForm.Items.Item("CntcCode").Click();
                        break;
                    case "CntcCode":
                        sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" + oDS_PS_MM010H.GetValue("U_CntcCode", 0).ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);

                        oDS_PS_MM010H.SetValue("U_CntcName", 0, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                        break;
                    case "Mat01":
                        if (oCol == "CGNo")
                        {
                            if ((oRow == oMat01.RowCount || oMat01.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat01.Columns.Item("CGNo").Cells.Item(oRow).Specific.Value.ToString().Trim()))
                            {
                                oMat01.FlushToDataSource();
                                PS_MM010_Add_MatrixRow(oMat01.RowCount, false);
                                oMat01.Columns.Item("CGNo").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }

                            sQry = "  SELECT  A.U_ItemCode, ";
                            sQry += "         A.U_ItemName,";
                            sQry += "         A.U_FrgnName,";
                            sQry += "         A.U_Qty,";
                            sQry += "         A.U_Weight,";
                            sQry += "         A.U_OrdType,";
                            sQry += "         A.U_Duedate,";
                            sQry += "         B.U_ItmBSort,";
                            sQry += "         B.U_ItmMSort,";
                            sQry += "         B.BuyUnitMsr,";
                            sQry += "         A.U_OutItmCd,";
                            sQry += "         A.U_OutSize,";
                            sQry += "         A.U_OutUnit,";
                            sQry += "         A.U_Auto,";
                            sQry += "         A.U_Comments,";
                            sQry += "         A.U_ProcCode,";
                            sQry += "         A.U_ProcName,";
                            sQry += "         A.U_OutCode,";
                            sQry += "         A.U_OutNote";
                            sQry += " FROM    [@PS_MM005H] A";
                            sQry += "         LEFT JOIN ";
                            sQry += "         [OITM] B ";
                            sQry += "             ON A.U_ItemCode = B.ItemCode";
                            sQry += " WHERE   A.U_CgNum = '" + oMat01.Columns.Item("CGNo").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);

                            if (oRecordSet01.RecordCount == 0)
                            {
                                errMessage = "조회 결과가 없습니다. 확인하세요";
                                throw new Exception();
                            }

                            oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_ItemCode").Value.ToString().Trim();
                            oMat01.Columns.Item("ItemName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_ItemName").Value.ToString().Trim();
                            oMat01.Columns.Item("Qty").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_Qty").Value.ToString().Trim();
                            oMat01.Columns.Item("Weight").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_Weight").Value.ToString().Trim();
                            oMat01.Columns.Item("ItemGpCd").Cells.Item(oRow).Specific.Select(oRecordSet01.Fields.Item("U_OrdType").Value.ToString().Trim());
                            oMat01.Columns.Item("DueDate").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_DueDate").Value.ToString("yyyyMMdd");
                            oMat01.Columns.Item("ItmBSort").Cells.Item(oRow).Specific.Select(oRecordSet01.Fields.Item("U_ItmBSort").Value.ToString().Trim());
                            oMat01.Columns.Item("OBasUnit").Cells.Item(oRow).Specific.Select(oRecordSet01.Fields.Item("BuyUnitMsr").Value.ToString().Trim());
                            oMat01.Columns.Item("OutSize").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_OutSize").Value.ToString().Trim();
                            oMat01.Columns.Item("OutUnit").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_OutUnit").Value.ToString().Trim();
                            oMat01.Columns.Item("Auto").Cells.Item(oRow).Specific.Select(oRecordSet01.Fields.Item("U_Auto").Value.ToString().Trim());
                            oMat01.Columns.Item("Comments").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_Comments").Value.ToString().Trim();
                            oMat01.Columns.Item("ProcCode").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_ProcCode").Value.ToString().Trim();
                            oMat01.Columns.Item("ProcName").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_ProcName").Value.ToString().Trim();

                            if (!string.IsNullOrEmpty(oRecordSet01.Fields.Item("U_OutCode").Value.ToString().Trim()))
                            {
                                oMat01.Columns.Item("OutCode").Cells.Item(oRow).Specific.Select(oRecordSet01.Fields.Item("U_OutCode").Value.ToString().Trim()); //외주사유코드
                            }
                            oMat01.Columns.Item("OutNote").Cells.Item(oRow).Specific.Value = oRecordSet01.Fields.Item("U_OutNote").Value.ToString().Trim(); //외주사유내용
                            oMat01.Columns.Item("CGNo").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                            for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                            {
                                if (!string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value))
                                {
                                    SumQty += Convert.ToInt32(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value);
                                }
                                SumWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value);
                            }

                            oMat01.AutoResizeColumns();
                            oForm.Items.Item("SumQty").Specific.Value = SumQty;
                            oForm.Items.Item("SumWeight").Specific.Value = SumWeight;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// Form Item Event
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">pVal</param>
        /// <param name="BubbleEvent">Bubble Event</param>
        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            switch (pVal.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                //    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                //    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                //    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
                //    Raise_EVENT_PICKER_CLICKED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                //    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_Drag: //39
                //    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                //    break;
            }
        }

        /// <summary>
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {

            int i;
            int Seq;
            string RequestNo;
            string DueDate;
            string Unit;
            string ItemName;
            string ItemCode;
            string Size;
            string RequestDate;
            string ItemType;
            string BPLId;
            string RFC_Sender = string.Empty;
            string sQry;
            string ItmBsort;
            string FileName;
            string Dir_Renamed;
            string CntcCode;
            string BEDNR;
            double Weight;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        //종전값을 넣기위해
                        oPurchase = oDS_PS_MM010H.GetValue("U_Purchase", 0).ToString().Trim();
                        oPQType = oDS_PS_MM010H.GetValue("U_PQType", 0).ToString().Trim();
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_MM010_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PS_MM010_MatrixSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                oMat01.FlushToDataSource();
                                if (oDS_PS_MM010H.GetValue("U_PQType", 0).ToString().Trim() == "10")
                                {
                                    oDS_PS_MM010H.SetValue("U_RFCAdms", 0, "Y");
                                    oDS_PS_MM010H.SetValue("U_RFCType", 0, "0");
                                    for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                                    {
                                        oDS_PS_MM010L.SetValue("U_GuBun", i, "0");
                                    }
                                }
                                else if (oDS_PS_MM010H.GetValue("U_PQType", 0).ToString().Trim() == "20")
                                {
                                    oDS_PS_MM010H.SetValue("U_RFCAdms", 0, "N");
                                }
                                oMat01.LoadFromDataSource();
                            }

                            //통합구매
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE && oDS_PS_MM010H.GetValue("U_PQType", 0).ToString().Trim() == "20" && oDS_PS_MM010H.GetValue("U_RFCAdms", 0).ToString().Trim() == "Y")
                            {
                                ProgressBar01.Text = "통합구매요청!";
                                //ErrNum = 101;
                                //oSeq = 1;
                                //oCount = 1;
                                Seq = 1;
                                oMat01.FlushToDataSource();

                                RequestDate = oDS_PS_MM010H.GetValue("U_DocDate", 0).ToString().Trim();
                                DueDate = oDS_PS_MM010H.GetValue("U_DocDate", 0).ToString().Trim();
                                BPLId = oDS_PS_MM010H.GetValue("U_BPLId", 0).ToString().Trim();

                                //담당자 성명(전화번호)
                                CntcCode = oDS_PS_MM010H.GetValue("U_CntcCode", 0).ToString().Trim();
                                sQry = "select U_FullName + '(' + right(Isnull(U_OffcTel,''),4) + ')' from [@PH_PY001A] where Code = '" + CntcCode + "'";
                                oRecordSet01.DoQuery(sQry);

                                BEDNR = oRecordSet01.Fields.Item(0).Value;

                                for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                                {
                                    if (string.IsNullOrEmpty(oDS_PS_MM010L.GetValue("U_E_BANFN", i).ToString().Trim()) && (string.IsNullOrEmpty(oDS_PS_MM010L.GetValue("U_E_BNFPO", i).ToString().Trim()) || oDS_PS_MM010L.GetValue("U_E_BNFPO", i).ToString().Trim() == "00000"))
                                    {
                                        ItemCode = oDS_PS_MM010L.GetValue("U_ItemCode", i).ToString().Trim();
                                        FileName = oDS_PS_MM010L.GetValue("U_FILENAME", i).ToString().Trim();
                                        Dir_Renamed = oDS_PS_MM010L.GetValue("U_Dir", i).ToString().Trim();
                                        ItmBsort = oDS_PS_MM010L.GetValue("U_ItemGpCd", i).ToString().Trim();


                                        //고정자산품의(60), 상품품의(50)
                                        if (oDS_PS_MM010H.GetValue("U_Purchase", 0).ToString().Trim() == "60" || oDS_PS_MM010H.GetValue("U_Purchase", 0).ToString().Trim() == "50")
                                        {
                                            ItemName = oDS_PS_MM010L.GetValue("U_ItemName", i).ToString().Trim();
                                            Size = oDS_PS_MM010L.GetValue("U_OutSize", i).ToString().Trim();
                                            Weight = Convert.ToDouble(oDS_PS_MM010L.GetValue("U_Weight", i));
                                            Unit = oDS_PS_MM010L.GetValue("U_OutUnit", i).ToString().Trim();
                                            ItemType = "";
                                            RequestNo = oDS_PS_MM010L.GetValue("U_CGNo", i).ToString().Trim();
                                        }
                                        else if (oDS_PS_MM010H.GetValue("U_Purchase", 0).ToString().Trim() == "40") //외주제작품의(40)
                                        {

                                            sQry = "Select U_Size, BuyUnitMsr, FrgnName, ItemName, U_ItmMSort From [OITM] Where ItemCode = '" + codeHelpClass.Left(ItemCode, 11) + "'";
                                            oRecordSet01.DoQuery(sQry);

                                            ItemName = oDS_PS_MM010L.GetValue("U_ItemName", i).ToString().Trim();
                                            Size = oDS_PS_MM010L.GetValue("U_OutSize", i).ToString().Trim();
                                            Weight = Convert.ToDouble(oDS_PS_MM010L.GetValue("U_Weight", i));
                                            Unit = oDS_PS_MM010L.GetValue("U_OutUnit", i).ToString().Trim();
                                            ItemType = oRecordSet01.Fields.Item("U_ItmMSort").Value.ToString().Trim();
                                            RequestNo = oDS_PS_MM010L.GetValue("U_CGNo", i).ToString().Trim();

                                        }
                                        else
                                        {
                                            sQry = "Select U_Size, BuyUnitMsr, FrgnName, ItemName, U_ItmMSort From [OITM] Where ItemCode = '" + ItemCode + "'";
                                            oRecordSet01.DoQuery(sQry);
                                            if (string.IsNullOrEmpty(oDS_PS_MM010L.GetValue("U_OutSize", i).ToString().Trim()))
                                            {
                                                Size = oRecordSet01.Fields.Item("U_Size").Value.ToString().Trim();
                                            }
                                            else
                                            {
                                                Size = oDS_PS_MM010L.GetValue("U_OutSize", i).ToString().Trim();
                                            }

                                            ItemName = oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim();
                                            Weight = Convert.ToDouble(oDS_PS_MM010L.GetValue("U_Weight", i));
                                            Unit = oRecordSet01.Fields.Item("BuyUnitMsr").Value.ToString().Trim();
                                            ItemType = oRecordSet01.Fields.Item("U_ItmMSort").Value.ToString().Trim();
                                            RequestNo = oDS_PS_MM010L.GetValue("U_CGNo", i).ToString().Trim();
                                            DueDate = oDS_PS_MM010L.GetValue("U_DueDate", i).ToString().Trim();
                                        }

                                        RFC_Sender = PS_RFC_Sender(BPLId, ItmBsort, ItemCode, ItemName, Size, Weight, Unit, RequestDate, DueDate, ItemType, RequestNo, BEDNR, i, oMat01.VisualRowCount - 2, FileName, Dir_Renamed).ToString().Trim();

                                        if (!string.IsNullOrEmpty(RFC_Sender))
                                        {
                                            oDS_PS_MM010L.SetValue("U_E_BANFN", i, codeHelpClass.Left(RFC_Sender, RFC_Sender.IndexOf("/")));
                                            oDS_PS_MM010L.SetValue("U_E_BNFPO", i, codeHelpClass.Right(RFC_Sender, RFC_Sender.Length - RFC_Sender.IndexOf("/") - 1));
                                            oDS_PS_MM010L.SetValue("U_MESSAGE", i, "");

                                            if (string.IsNullOrEmpty(codeHelpClass.Left(RFC_Sender, RFC_Sender.IndexOf("/") - 1)) || string.IsNullOrEmpty(codeHelpClass.Right(RFC_Sender, RFC_Sender.Length - RFC_Sender.IndexOf("/"))))
                                            {
                                                Seq = 0;
                                            }
                                            oDS_PS_MM010L.SetValue("U_GuBun", i, Convert.ToString(Seq));
                                            Seq = 1;
                                        }

                                        ProgressBar01.Value += 1;
                                        ProgressBar01.Text = ProgressBar01.Value + "/" + Convert.ToString(oMat01.VisualRowCount - 2 + 1) + "건 처리중...!";
                                    }
                                }
                                oMat01.LoadFromDataSource();
                                oDS_PS_MM010H.SetValue("U_RFCType", 0, Convert.ToString(Seq));
                            }
                            PS_MM010_Delete_EmptyRow();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                        {
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        //종전값을 넣기위해 이값에 입력시 화면에 값이 이미 사라져서 변수에 입력을 못함. BeforeAction = True에 입력함
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            PS_MM010_FormItemEnabled();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
                        {
                            PS_MM010_Initialization();
                            PS_MM010_FormItemEnabled();
                            PS_MM010_FormClear();
                            oDS_PS_MM010H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                            PS_MM010_Add_MatrixRow(0, true);
                            oForm.Items.Item("SumQty").Specific.Value = 0;
                            oForm.Items.Item("SumWeight").Specific.Value = 0;
                        }
                    }
                    else if (pVal.ItemUID == "Btn01")
                    {
                        if (oDS_PS_MM010H.GetValue("U_PQType", 0).ToString().Trim() == "10") //견적형태:자체구매(10)
                        {
                            if (oDS_PS_MM010H.GetValue("U_Purchase", 0).ToString().Trim() != "30" || oDS_PS_MM010H.GetValue("U_Purchase", 0).ToString().Trim() != "40") //품의구분:가공비품의(30), 외주제작품의(40)
                            {
                                System.Threading.Thread thread = new System.Threading.Thread(PS_MM010_Print_Report01);
                                thread.SetApartmentState(System.Threading.ApartmentState.STA);
                                thread.Start();
                            }
                        }
                        else if (oDS_PS_MM010H.GetValue("U_PQType", 0).ToString().Trim() == "20") //견적형태:통합구매(20)
                        {
                            System.Threading.Thread thread = new System.Threading.Thread(PS_MM010_Print_Report02);
                            thread.SetApartmentState(System.Threading.ApartmentState.STA);
                            thread.Start();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// KEY_DOWN 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int pos;
            string sFile;
            string FileName;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            FileListBoxForm fileListBoxForm = new FileListBoxForm();

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "CntcCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "CGNo")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("CGNo").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
                            }
                            if (pVal.ColUID == "FILENAME")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("FILENAME").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    sFile = fileListBoxForm.OpenDialog(fileListBoxForm, "*.*", "파일선택", "C:\\");
                                    if (string.IsNullOrEmpty(sFile))
                                    {
                                        PSH_Globals.SBO_Application.StatusBar.SetText("파일을 선택해 주세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        return;
                                    }
                                    else
                                    {
                                        oMat01.Columns.Item("Dir").Cells.Item(pVal.Row).Specific.Value = sFile;
                                        pos = sFile.LastIndexOf("\\");

                                        FileName = codeHelpClass.Mid(sFile, pos + 1, sFile.Length - pos - 1);

                                        oMat01.Columns.Item("FILENAME").Cells.Item(pVal.Row).Specific.Value = FileName;
                                        oMat01.AutoResizeColumns();
                                    }
                                }
                            }
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.CharPressed == 38) //방향키(↑)
                        {
                            if (pVal.Row > 1 && pVal.Row <= oMat01.VisualRowCount)
                            {
                                oForm.Freeze(true);
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row - 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oForm.Freeze(false);
                            }
                        }
                        else if (pVal.CharPressed == 40) //방향키(↓)
                        {
                            if (pVal.Row > 0 && pVal.Row < oMat01.VisualRowCount)
                            {
                                oForm.Freeze(true);
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oForm.Freeze(false);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.ItemUID == "Mat01")
                {
                    if (pVal.Row > 0)
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = pVal.ColUID;
                        oLastColRow01 = pVal.Row;
                    }
                }
                else
                {
                    oLastItemUID01 = pVal.ItemUID;
                    oLastColUID01 = "";
                    oLastColRow01 = 0;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }

        /// <summary>
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Purchase" || pVal.ItemUID == "BPLId")
                    {
                        oMat01.Clear();
                        oDS_PS_MM010L.Clear();
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PS_MM010_Add_MatrixRow(0, false);
                        }

                        if (oForm.Items.Item("Purchase").Specific.Value.ToString().Trim() == "40")
                        {
                            oMat01.Columns.Item("ItemName").Editable = true;
                            oMat01.Columns.Item("OutSize").Editable = true;
                            oMat01.Columns.Item("OutUnit").Editable = true;
                        }
                        else
                        {
                            oMat01.Columns.Item("ItemName").Editable = false;
                            oMat01.Columns.Item("OutSize").Editable = false;
                            oMat01.Columns.Item("OutUnit").Editable = false;
                        }
                        PS_MM010_FlushToItemValue(pVal.ItemUID, 0, "");
                    }
                    else if (pVal.ItemUID == "PQType")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (oForm.Items.Item("PQType").Specific.Value == "20")
                            {
                                oForm.Items.Item("RFCAdms").Specific.Select("N");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;

                            oMat01.SelectRow(pVal.Row, true, false);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }

        /// <summary>
        /// MATRIX_LINK_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "E_BANFN")
                        {
                            //통합구매이며 통합구매요청번호가 있을때
                            if (oForm.Items.Item("PQType").Specific.Value.ToString().Trim() == "20" && !string.IsNullOrEmpty(oMat01.Columns.Item("E_BANFN").Cells.Item(pVal.Row).Specific.String))
                            {
                                PSH_Globals.SBO_Application.MessageBox("Migration 진행중입니다. 진행되는 즉시 패치하겠습니다.");
                                //TempForm01 = new PS_FTP();
                                //TempForm01.LoadForm(oMat01.Columns.Item("CGNo").Cells.Item(pVal.Row).Specific.String, oMat01.Columns.Item("E_BANFN").Cells.Item(pVal.Row).Specific.String, oMat01.Columns.Item("E_BNFPO").Cells.Item(pVal.Row).Specific.String);
                                //BubbleEvent = false;
                            }
                            else
                            {
                            }
                        }

                        else if (pVal.ColUID == "CGNo")
                        {
                            if (!string.IsNullOrEmpty(oMat01.Columns.Item("CGNo").Cells.Item(pVal.Row).Specific.String))
                            {
                                PS_MM010_S PS_MM010_S = new PS_MM010_S();
                                PS_MM010_S.LoadForm(oMat01.Columns.Item("CGNo").Cells.Item(pVal.Row).Specific.String);
                            }
                            else
                            {
                            }
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }

        /// <summary>
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int Qty;
            string ItemCode;
            double Calculate_Weight;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "CntcCode")
                        {
                            PS_MM010_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "CGNo")
                            {
                                PS_MM010_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "Qty")
                            {
                                oMat01.FlushToDataSource();
                                ItemCode = oDS_PS_MM010L.GetValue("U_ItemCode", pVal.Row - 1).ToString().Trim();
                                Qty = Convert.ToInt32(oDS_PS_MM010L.GetValue("U_Qty", pVal.Row - 1));

                                Calculate_Weight = dataHelpClass.Calculate_Weight(ItemCode, Qty, oForm.Items.Item("BPLId").Specific.Value.ToString().Trim());

                                oDS_PS_MM010L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(Calculate_Weight));
                                oMat01.LoadFromDataSource();
                                oMat01.Columns.Item("Weight").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                BubbleEvent = false;
            }
            finally
            {
            }
        }

        /// <summary>
        /// MATRIX_LOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;
            int SumQty = 0;
            double SumWeight = 0;

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        if (!string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value))
                        {
                            SumQty += Convert.ToInt32(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value);
                        }
                        SumWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value);
                    }

                    oForm.Items.Item("SumQty").Specific.Value = SumQty;
                    oForm.Items.Item("SumWeight").Specific.Value = SumWeight;

                    oMat01.AutoResizeColumns();
                    PS_MM010_Add_MatrixRow(oMat01.RowCount, false);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }

        /// <summary>
        /// FORM_UNLOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_UNLOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    SubMain.Remove_Forms(oFormUniqueID);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM010H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM010L);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }

        /// <summary>
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            int i;
            string sQry;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                            {
                                sQry = "Select b.DocNum From [@PS_MM030L] a Inner Join [@PS_MM030H] b On a.DocEntry = b.DocEntry ";
                                sQry += "where a.U_PQDocNum = '" + oForm.Items.Item("DocNum").Specific.Value.ToString().Trim() + "' ";
                                sQry += "And a.U_PQLinNum = '" + oMat01.Columns.Item("LineNum").Cells.Item(i + 1).Specific.Value.ToString().Trim() + "' And b.Canceled = 'N'";
                                oRecordSet01.DoQuery(sQry);

                                if (oRecordSet01.Fields.Item(0).Value == 0 || string.IsNullOrEmpty(oRecordSet01.Fields.Item(0).Value))
                                {
                                }
                                else
                                {
                                    errMessage = "" + i + 1 + "번 라인의 구매견적이 발주등록 되었습니다. 취소할 수 없습니다.";
                                    throw new Exception();
                                }
                            }
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            PS_MM010_FormItemEnabled();
                            oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            if (oMat01.RowCount != oMat01.VisualRowCount)
                            {
                                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    oMat01.Columns.Item("LineNum").Cells.Item(i + 1).Specific.Value = i + 1;
                                }
                                oMat01.FlushToDataSource();
                                oDS_PS_MM010L.RemoveRecord(oDS_PS_MM010L.Size - 1); // Mat01에 마지막라인(빈라인) 삭제
                                oMat01.Clear();
                                oMat01.LoadFromDataSource();
                            }
                            break;
                        case "1281": //찾기
                            PS_MM010_FormItemEnabled();
                            oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("SumQty").Specific.Value = 0;
                            oForm.Items.Item("SumWeight").Specific.Value = 0;

                            //아이디별 사업장 세팅
                            oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                            //아이디별 사번 세팅
                            //수퍼유저인 경우는 사번 미표기(2016.01.08 송명규)
                            if (dataHelpClass.User_SuperUserYN() == "N")
                            {
                                oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
                            }
                            break;
                        case "1282": //추가
                            PS_MM010_Initialization();
                            PS_MM010_FormItemEnabled();
                            PS_MM010_FormClear();
                            oDS_PS_MM010H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                            PS_MM010_Add_MatrixRow(0, true);
                            oForm.Items.Item("SumQty").Specific.Value = 0;
                            oForm.Items.Item("SumWeight").Specific.Value = 0;
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
                            PS_MM010_FormItemEnabled();
                            PS_MM010_FlushToItemValue("BPLId", 0, "");

                            //if (oMat01.VisualRowCount > 0)
                            //{
                            //    if (!string.IsNullOrEmpty(oMat01.Columns.Item("CGNo").Cells.Item(oMat01.VisualRowCount).Specific.Value))
                            //    {
                            //        if (oDS_PS_MM010H.GetValue("Status", 0) == "O")
                            //        {
                            //            PS_MM010_Add_MatrixRow(oMat01.RowCount, false);
                            //        }
                            //    }
                            //}
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// FormDataEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }

        /// <summary>
        /// RightClickEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                }

                switch (pVal.ItemUID)
                {
                    case "Mat01":
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
            }
        }
    }
}
