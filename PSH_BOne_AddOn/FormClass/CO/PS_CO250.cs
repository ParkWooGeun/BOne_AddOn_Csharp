using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 개인별 퇴충계산
    /// </summary>
    internal class PS_CO250 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_CO250H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_CO250L; //등록라인
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        public override void LoadForm()
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO250.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_CO250_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_CO250");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry"; //Code로 지정하면 레코드 버튼이 이동이 순차적으로 동작하지 않음

                oForm.Freeze(true);
                CreateItems();
                ComboBox_Setting();
                FormItemEnabled();

                oForm.EnableMenu("1283", true); //삭제
                oForm.EnableMenu("1287", true); //복제
                oForm.EnableMenu("1286", false); //닫기
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", true); //행삭제

                oForm.Items.Item("DocEntry").Visible = false; //레코드 이동 버튼의 순차 동작을 위해 추가한 DocEntry의 Visible을 false로 지정
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void CreateItems()
        {
            try
            {
                oDS_PS_CO250H = oForm.DataSources.DBDataSources.Item("@PS_CO250H");
                oDS_PS_CO250L = oForm.DataSources.DBDataSources.Item("@PS_CO250L");

                oMat01 = oForm.Items.Item("Mat01").Specific;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 콤보박스 세팅
        /// </summary>
        private void ComboBox_Setting()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
                oRecordSet01.DoQuery(sQry);

                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                oDS_PS_CO250H.SetValue("U_BPLId", 0, dataHelpClass.User_BPLID());
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void FormItemEnabled()
        {
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("Code").Enabled = false;

                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("Code").Enabled = true;

                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("Code").Enabled = false;

                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 행추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void Add_MatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                if (RowIserted == false)
                {
                    oDS_PS_CO250L.InsertRecord(oRow);
                }

                oMat01.AddRow();
                oDS_PS_CO250L.Offset = oRow;
                oDS_PS_CO250L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 필수입력사항 체크(Header)
        /// </summary>
        /// <returns></returns>
        private bool HeaderSpaceLineDel()
        {
            bool returnValue = false;
            string errCode = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_CO250H.GetValue("U_DocDate", 0)))
                {
                    errCode = "1";
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oDS_PS_CO250H.GetValue("U_BPLId", 0)))
                {
                    errCode = "2";
                    throw new Exception();
                }

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("기준일자는 필수입력사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장은 필수입력사항입니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }

            return returnValue;
        }

        /// <summary>
        /// 필수입력사항 체크(Line)
        /// </summary>
        /// <returns></returns>
        private bool MatrixSpaceLineDel()
        {
            bool returnValue = false;
            int i;
            string errCode = string.Empty;

            try
            {
                oMat01.FlushToDataSource();

                if (oMat01.VisualRowCount == 0)
                {
                    errCode = "1";
                    throw new Exception();
                }

                for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                {

                }

                oMat01.LoadFromDataSource();
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("라인 데이터가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// Matrix 마지막 빈행 삭제
        /// </summary>
        private void Delete_EmptyRow()
        {
            try
            {
                oMat01.FlushToDataSource();

                for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oDS_PS_CO250L.GetValue("U_CycleCod", i).ToString().Trim()))
                    {
                        oDS_PS_CO250L.RemoveRecord(i);
                    }
                }

                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        private void LoadData()
        {

            int i = 0;
            string sQry = string.Empty;
            string DocDate = string.Empty;
            string BPLId = string.Empty;
            string errCode = string.Empty;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgBar01 = null;

            try
            {
                DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
                BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();

                sQry = "  SELECT    MSTCOD,";
                sQry += "           FullName,";
                sQry += "           TeamCode,";
                sQry += "           TeamName,";
                sQry += "           RspCode,";
                sQry += "           RspName,";
                sQry += "           ClsCode,";
                sQry += "           ClsName,";
                sQry += "           InpDat,";
                sQry += "           GrpDat,";
                sQry += "           RETDAT,";
                sQry += "           JIGTYP,";
                sQry += "           JIGTYPNM,";
                sQry += "           JIGCOD,";
                sQry += "           JIGCODNM,";
                sQry += "           PAYTYP,";
                sQry += "           PAYTYPNM,";
                sQry += "           YYCnt,";
                sQry += "           MMCnt,";
                sQry += "           DDCnt,";
                sQry += "           MonthCnt,";
                sQry += "           BAESU,";
                sQry += "           PAY1,";
                sQry += "           PAY2,";
                sQry += "           PAY3,";
                sQry += "           BNSTOT,";
                sQry += "           YUNSU,";
                sQry += "           HUGA,";
                sQry += "           AVGPAY,";
                sQry += "           AVGBNS,";
                sQry += "           AVGYUNSU,";
                sQry += "           AVGHUGA,";
                sQry += "           AVGTOT,";
                sQry += "           ToiAmt,";
                sQry += "           BirthDat";
                sQry += " FROM      ZPS_CO250L";
                sQry += " WHERE     CLTCOD = '" + BPLId + "'";
                sQry += "           AND DocDate = '" + DocDate + "'";
                sQry += " ORDER BY  TeamCode,";
                sQry += "           RspCode,";
                sQry += "           ClsCode,";
                sQry += "           JIGCOD ";
                oRecordSet01.DoQuery(sQry);

                oForm.Freeze(true);

                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

                oMat01.Clear();
                oDS_PS_CO250L.Clear();

                if (oRecordSet01.RecordCount == 0)
                {
                    errCode = "1";
                    throw new Exception();
                }
                
                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_CO250L.Size)
                    {
                        oDS_PS_CO250L.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PS_CO250L.Offset = i;
                    oDS_PS_CO250L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_CO250L.SetValue("U_MSTCOD", i, oRecordSet01.Fields.Item("MSTCOD").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_FullName", i, oRecordSet01.Fields.Item("FullName").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_BirthDat", i, oRecordSet01.Fields.Item("BirthDat").Value.ToString("yyyyMMdd"));
                    oDS_PS_CO250L.SetValue("U_TeamCode", i, oRecordSet01.Fields.Item("TeamCode").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_TeamName", i, oRecordSet01.Fields.Item("TeamName").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_RspCode", i, oRecordSet01.Fields.Item("RspCode").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_RspName", i, oRecordSet01.Fields.Item("RspName").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_ClsCode", i, oRecordSet01.Fields.Item("ClsCode").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_ClsName", i, oRecordSet01.Fields.Item("ClsName").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_InpDat", i, oRecordSet01.Fields.Item("InpDat").Value.ToString("yyyyMMdd"));
                    oDS_PS_CO250L.SetValue("U_GrpDat", i, oRecordSet01.Fields.Item("GrpDat").Value.ToString("yyyyMMdd"));
                    oDS_PS_CO250L.SetValue("U_RETDAT", i, oRecordSet01.Fields.Item("RETDAT").Value.ToString("yyyyMMdd"));
                    oDS_PS_CO250L.SetValue("U_JIGTYP", i, oRecordSet01.Fields.Item("JIGTYP").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_JIGTYPNM", i, oRecordSet01.Fields.Item("JIGTYPNM").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_JIGCOD", i, oRecordSet01.Fields.Item("JIGCOD").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_JIGCODNM", i, oRecordSet01.Fields.Item("JIGCODNM").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_PAYTYP", i, oRecordSet01.Fields.Item("PAYTYP").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_PAYTYPNM", i, oRecordSet01.Fields.Item("PAYTYPNM").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_YYCnt", i, oRecordSet01.Fields.Item("YYCnt").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_MMCnt", i, oRecordSet01.Fields.Item("MMCnt").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_DDCnt", i, oRecordSet01.Fields.Item("DDCnt").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_MonthCnt", i, oRecordSet01.Fields.Item("MonthCnt").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_BAESU", i, oRecordSet01.Fields.Item("BAESU").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_PAY1", i, oRecordSet01.Fields.Item("PAY1").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_PAY2", i, oRecordSet01.Fields.Item("PAY2").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_PAY3", i, oRecordSet01.Fields.Item("PAY3").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_BNSTOT", i, oRecordSet01.Fields.Item("BNSTOT").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_YUNSU", i, oRecordSet01.Fields.Item("YUNSU").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_HUGA", i, oRecordSet01.Fields.Item("HUGA").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_AVGPAY", i, oRecordSet01.Fields.Item("AVGPAY").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_AVGBNS", i, oRecordSet01.Fields.Item("AVGBNS").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_AVGYUNSU", i, oRecordSet01.Fields.Item("AVGYUNSU").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_AVGHUGA", i, oRecordSet01.Fields.Item("AVGHUGA").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_AVGTOT", i, oRecordSet01.Fields.Item("AVGTOT").Value.ToString().Trim());
                    oDS_PS_CO250L.SetValue("U_ToiAmt", i, oRecordSet01.Fields.Item("ToiAmt").Value.ToString().Trim());

                    oRecordSet01.MoveNext();
                    ProgBar01.Value = ProgBar01.Value + 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("조회 결과가 없습니다.확인하세요.");
                }
                else 
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                ProgBar01.Stop();
                oForm.Freeze(false);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
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
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    //Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    //Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                    //Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                    //Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    //Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    //Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                    //Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    //Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_Drag: //39
                    //Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                    break;
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (MatrixSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            string Code = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_CO250'", "");

                            oDS_PS_CO250H.SetValue("Code", 0, Code);
                            oDS_PS_CO250H.SetValue("Name", 0, Code);
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PSH_Globals.SBO_Application.ActivateMenuItem("1282");
                        }
                    }
                    else if (pVal.ItemUID == "Btn01")
                    {
                        if (HeaderSpaceLineDel() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        LoadData();
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    FormItemEnabled();
                    Add_MatrixRow(oMat01.RowCount, false);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO250H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_CO250L);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            break;
                        case "1281": //찾기
                            FormItemEnabled();
                            break;
                        case "1282": //추가
                            FormItemEnabled();
                            Add_MatrixRow(0, true);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            FormItemEnabled();
                            break;
                        case "1287":// 복제
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                if (BusinessObjectInfo.BeforeAction == true)
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
                else if (BusinessObjectInfo.BeforeAction == false)
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
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
