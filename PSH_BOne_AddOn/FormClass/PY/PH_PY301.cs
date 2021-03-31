using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 학자금신청등록
    /// </summary>
    internal class PH_PY301 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat1;
        private SAPbouiCOM.DBDataSource oDS_PH_PY301A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY301B;
        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY301.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY301_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY301");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PH_PY301_CreateItems();
                PH_PY301_EnableMenus();
                PH_PY301_SetDocument(oFormDocEntry01);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("LoadForm_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        /// <returns></returns>
        private void PH_PY301_CreateItems()
        {
            string sQry;
           
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                oDS_PH_PY301A = oForm.DataSources.DBDataSources.Item("@PH_PY301A");
                oDS_PH_PY301B = oForm.DataSources.DBDataSources.Item("@PH_PY301B");

                oMat1 = oForm.Items.Item("Mat01").Specific;

                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                oMat1.AutoResizeColumns();

                //사업장
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                //분기
                oForm.Items.Item("Quarter").Specific.ValidValues.Add("", "");
                oForm.Items.Item("Quarter").Specific.ValidValues.Add("01", "1/4 혹은 1학기");
                oForm.Items.Item("Quarter").Specific.ValidValues.Add("02", "2/4");
                oForm.Items.Item("Quarter").Specific.ValidValues.Add("03", "3/4 혹은 2학기");
                oForm.Items.Item("Quarter").Specific.ValidValues.Add("04", "4/4");
                oForm.Items.Item("Quarter").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("Quarter").DisplayDesc = true;

                //매트릭스-성별
                oMat1.Columns.Item("Sex").ValidValues.Add("", "");
                oMat1.Columns.Item("Sex").ValidValues.Add("01", "남자");
                oMat1.Columns.Item("Sex").ValidValues.Add("02", "여자");
                oMat1.Columns.Item("Sex").DisplayDesc = true;

                //매트릭스-학교
                oMat1.Columns.Item("SchCls").ValidValues.Add("", "");
                sQry = "  SELECT    T1.U_Code,";
                sQry += "           T1.U_CodeNm";
                sQry += " FROM      [@PS_HR200H] AS T0";
                sQry += "           INNER JOIN";
                sQry += "           [@PS_HR200L] AS T1";
                sQry += "               ON T0.Code = T1.Code";
                sQry += " WHERE     T0.Code = 'P222'";
                sQry += "           AND T1.U_UseYN = 'Y'";
                sQry += " ORDER BY  T1.U_Seq";

                dataHelpClass.GP_MatrixSetMatComboList(oMat1.Columns.Item("SchCls"), sQry, "", "");

                //매트릭스-학년
                oMat1.Columns.Item("Grade").ValidValues.Add("", "");
                oMat1.Columns.Item("Grade").ValidValues.Add("01", "1학년");
                oMat1.Columns.Item("Grade").ValidValues.Add("02", "2학년");
                oMat1.Columns.Item("Grade").ValidValues.Add("03", "3학년");
                oMat1.Columns.Item("Grade").ValidValues.Add("04", "4학년");
                oMat1.Columns.Item("Grade").ValidValues.Add("05", "5학년");
                oMat1.Columns.Item("Grade").DisplayDesc = true;

                //매트릭스-회차
                oMat1.Columns.Item("Count").ValidValues.Add("", "");
                oMat1.Columns.Item("Count").ValidValues.Add("01", "1차");
                oMat1.Columns.Item("Count").ValidValues.Add("02", "2차");
                oMat1.Columns.Item("Count").DisplayDesc = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);  //메모리 해제
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY301_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", false); // 삭제
                oForm.EnableMenu("1287", true); // 복제
                oForm.EnableMenu("1284", true); // 취소
                oForm.EnableMenu("1293", true); // 행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY301_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry01))
                {
                    PH_PY301_FormItemEnabled();
                    PH_PY301_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY301_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY301_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY301_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("StdYear").Enabled = true;
                    oForm.Items.Item("Quarter").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    
                    PH_PY301_FormClear(); //폼 DocEntry 세팅

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅
                    
                    oDS_PH_PY301A.SetValue("U_StdYear", 0, DateTime.Now.ToString("yyyy")); //년도 세팅

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("StdYear").Enabled = true;
                    oForm.Items.Item("Quarter").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = true;
                    
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("CntcCode").Enabled = false;
                    oForm.Items.Item("StdYear").Enabled = false;
                    oForm.Items.Item("Quarter").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false); //접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Matirx 행 추가
        /// </summary>
        private void PH_PY301_AddMatrixRow()
        {
            int oRow;

            try
            {
                oForm.Freeze(true);
                
                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;
                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY301B.GetValue("U_Name", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY301B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY301B.InsertRecord((oRow));
                        }
                        oDS_PH_PY301B.Offset = oRow;
                        oDS_PH_PY301B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY301B.SetValue("U_Name", oRow, "");
                        oDS_PH_PY301B.SetValue("U_GovID", oRow, "");
                        oDS_PH_PY301B.SetValue("U_Sex", oRow, "");
                        oDS_PH_PY301B.SetValue("U_SchCls", oRow, "");
                        oDS_PH_PY301B.SetValue("U_SchName", oRow, "");
                        oDS_PH_PY301B.SetValue("U_Grade", oRow, "");
                        oDS_PH_PY301B.SetValue("U_EntFee", oRow, "0");
                        oDS_PH_PY301B.SetValue("U_Tuition", oRow, "0");
                        oDS_PH_PY301B.SetValue("U_Count", oRow, "");
                        oDS_PH_PY301B.SetValue("U_PayCnt", oRow, "");
                        oDS_PH_PY301B.SetValue("U_PayYN", oRow, "");
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY301B.Offset = oRow - 1;
                        oDS_PH_PY301B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY301B.SetValue("U_Name", oRow - 1, "");
                        oDS_PH_PY301B.SetValue("U_GovID", oRow - 1, "");
                        oDS_PH_PY301B.SetValue("U_Sex", oRow - 1, "");
                        oDS_PH_PY301B.SetValue("U_SchCls", oRow - 1, "");
                        oDS_PH_PY301B.SetValue("U_SchName", oRow - 1, "");
                        oDS_PH_PY301B.SetValue("U_Grade", oRow - 1, "");
                        oDS_PH_PY301B.SetValue("U_EntFee", oRow - 1, "0");
                        oDS_PH_PY301B.SetValue("U_Tuition", oRow - 1, "0");
                        oDS_PH_PY301B.SetValue("U_Count", oRow - 1, "");
                        oDS_PH_PY301B.SetValue("U_PayCnt", oRow, "");
                        oDS_PH_PY301B.SetValue("U_PayYN", oRow - 1, "");
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY301B.Offset = oRow;
                    oDS_PH_PY301B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY301B.SetValue("U_Name", oRow, "");
                    oDS_PH_PY301B.SetValue("U_GovID", oRow, "");
                    oDS_PH_PY301B.SetValue("U_Sex", oRow, "");
                    oDS_PH_PY301B.SetValue("U_SchCls", oRow, "");
                    oDS_PH_PY301B.SetValue("U_SchName", oRow, "");
                    oDS_PH_PY301B.SetValue("U_Grade", oRow, "");
                    oDS_PH_PY301B.SetValue("U_EntFee", oRow, "0");
                    oDS_PH_PY301B.SetValue("U_Tuition", oRow, "0");
                    oDS_PH_PY301B.SetValue("U_Count", oRow, "");
                    oDS_PH_PY301B.SetValue("U_PayCnt", oRow, "");
                    oDS_PH_PY301B.SetValue("U_PayYN", oRow, "");
                    oMat1.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY301_AddMatrixRow_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PH_PY301_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY301'", "");
                if (Convert.ToDouble(DocEntry) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY301_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY301_DataValidCheck()
        {
            bool functionReturnValue = false;
            int i;
            string CLTCOD;
            string StdYear;
            string Quarter;
            string Count;
            string sQry;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            
            try
            {
                //사업장
                if (string.IsNullOrEmpty(oDS_PH_PY301A.GetValue("U_CLTCOD", 0)))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }
                //년도
                if (string.IsNullOrEmpty(oDS_PH_PY301A.GetValue("U_StdYear", 0)))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("년도는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("StdYear").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }
                //사번
                if (string.IsNullOrEmpty(oDS_PH_PY301A.GetValue("U_CntcCode", 0)))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사번은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CntcCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }
                //분기
                if (string.IsNullOrEmpty(oDS_PH_PY301A.GetValue("U_Quarter", 0)))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("분기는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("Quarter").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                CLTCOD = oDS_PH_PY301A.GetValue("U_CLTCOD", 0);
                StdYear = oDS_PH_PY301A.GetValue("U_StdYear", 0);
                Quarter = oDS_PH_PY301A.GetValue("U_Quarter", 0);

                //라인
                if (oMat1.VisualRowCount > 1)
                {
                    for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
                    {
                        //학교
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("SchCls").Cells.Item(i).Specific.Value))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("학교는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("SchCls").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }
                        //학교명
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("SchName").Cells.Item(i).Specific.Value))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("학교명은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("SchName").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }
                        //학년
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("Grade").Cells.Item(i).Specific.Value))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("학년은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("Grade").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }
                        //회차
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("Count").Cells.Item(i).Specific.Value))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("회차는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("Count").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }
                        Count = oMat1.Columns.Item("Count").Cells.Item(i).Specific.Value;

                        sQry = "Select Cnt = Count(*) From [@PH_PY301A] a Inner Join [@PH_PY301B] b On a.DocEntry = b.DocEntry and a.Canceled = 'N' ";
                        sQry += " Where a.U_CLTCOD = '" + CLTCOD + "' And a.U_StdYear = '" + StdYear + "' and a.U_Quarter = '" + Quarter + "' ";
                        sQry += " And b.U_Count = '" + Count + "' and b.U_PayYN = 'Y'";

                        oRecordSet.DoQuery(sQry);

                        if (oRecordSet.Fields.Item(0).Value > 0)
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("지급완료처리가 되어 추가/수정을 할 수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            return functionReturnValue;
                        }
                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("라인 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    return functionReturnValue;
                }

                oMat1.FlushToDataSource();
                if (oDS_PH_PY301B.Size > 1)
                {
                    oDS_PH_PY301B.RemoveRecord(oDS_PH_PY301B.Size - 1);
                }
                oMat1.LoadFromDataSource();

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 지급횟수 계산
        /// </summary>
        /// <param name="pGovID"></param>
        /// <param name="pSchCls"></param>
        /// <param name="pDocEntry"></param>
        /// <returns></returns>
        private int PH_PY301_GetPayCount(string pGovID, string pSchCls, short pDocEntry)
        {
            int functionReturnValue = 0;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "EXEC PH_PY301_01 '" + pGovID + "','" + pSchCls + "','" + pDocEntry + "'";
                oRecordSet.DoQuery(sQry);
                functionReturnValue = oRecordSet.Fields.Item("PayCount").Value;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_CheckAmt_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return functionReturnValue;
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
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY301_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY301_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY301_FormItemEnabled();
                                PH_PY301_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY301_FormItemEnabled();
                                PH_PY301_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY301_FormItemEnabled();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Mat01":
                            if (pVal.Row > 0)
                            {
                                oMat1.SelectRow(pVal.Row, true, false);
                            }
                            break;
                    }

                    switch (pVal.ItemUID)
                    {
                        case "Mat01":
                            if (pVal.Row > 0)
                            {
                                oLastItemUID = pVal.ItemUID;
                                oLastColUID = pVal.ColUID;
                                oLastColRow = pVal.Row;
                            }
                            break;
                        default:
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = "";
                            oLastColRow = 0;
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string GovID;
            string GovID1;
            string SchCls;
            string Sex = string.Empty;
            string GovID2;

            short loopCount;
            double FeeTot = 0;
            double TuiTot = 0;
            double Total;
            double PreTuition;
            double Tuition;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "CntcCode":
                                oDS_PH_PY301A.SetValue("U_CntcName", 0, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'",""));
                                break;

                            case "Mat01":

                                if (pVal.ColUID == "Name")
                                {

                                    oMat1.FlushToDataSource();

                                    GovID = dataHelpClass.Get_ReData( "U_FamPer",  "U_FamNam",  "[@PH_PY001D]",  "'" + oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'",  " AND Code = '" + oDS_PH_PY301A.GetValue("U_CntcCode", 0) + "'");

                                    if (GovID.Substring(6,1) == "1" || GovID.Substring(6, 1) == "3" || GovID.Substring(6, 1) == "5")
                                    {
                                        Sex = "01";
                                    }
                                    else if (GovID.Substring(6, 1) == "2" || GovID.Substring(6, 1) == "4" || GovID.Substring(6, 1) == "6")
                                    {
                                        Sex = "02";
                                    }

                                    GovID1 = GovID.Substring(0, 6);
                                    GovID2 = GovID.Substring(6, 7);
                                    GovID = GovID1 + "-" + GovID2;

                                    oDS_PH_PY301B.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PH_PY301B.SetValue("U_GovID", pVal.Row - 1, GovID); //주민등록번호
                                    oDS_PH_PY301B.SetValue("U_Sex", pVal.Row - 1, Sex); //성별
                                    oDS_PH_PY301B.SetValue("U_PayYN", pVal.Row - 1, "N"); //지급완료여부

                                    oMat1.LoadFromDataSource();

                                    PH_PY301_AddMatrixRow();

                                    
                                }
                                else if (pVal.ColUID == "EntFee") //입학금 입력 시
                                {
                                    oMat1.FlushToDataSource();

                                    //학교선택을 하지 않으면 에러 메시지 출력
                                    if (string.IsNullOrEmpty(oDS_PH_PY301B.GetValue("U_SchCls", pVal.Row - 1)))
                                    {
                                        dataHelpClass.MDC_GF_Message("학교를 먼저 선택하십시오.", "E");
                                        oDS_PH_PY301B.SetValue("U_EntFee", pVal.Row - 1, "0");
                                        BubbleEvent = false;
                                    }

                                    //입학금 합계 계산
                                    for (loopCount = 0; loopCount <= oMat1.RowCount - 1; loopCount++)
                                    {
                                        FeeTot += Convert.ToDouble(oDS_PH_PY301B.GetValue("U_EntFee", loopCount));
                                    }
                                    oMat1.LoadFromDataSource();

                                    oDS_PH_PY301A.SetValue("U_FeeTot", 0, Convert.ToString(FeeTot));

                                    TuiTot = Convert.ToDouble(oDS_PH_PY301A.GetValue("U_TuiTot", 0));
                                    Total = FeeTot + TuiTot;

                                    oDS_PH_PY301A.SetValue("U_Total", 0, Convert.ToString(Total));
                                }
                                else if (pVal.ColUID == "Tuition") //등록금 입력 시
                                {
                                    PreTuition = Convert.ToDouble(oDS_PH_PY301B.GetValue("U_Tuition", pVal.Row - 1));

                                    oMat1.FlushToDataSource();

                                    //학교선택을 하지 않으면 에러 메시지 출력
                                    if (string.IsNullOrEmpty(oDS_PH_PY301B.GetValue("U_SchCls", pVal.Row - 1)))
                                    {
                                        dataHelpClass.MDC_GF_Message( "학교를 먼저 선택하십시오.",  "E");
                                        oDS_PH_PY301B.SetValue("U_Tuition", pVal.Row - 1, "0");
                                        BubbleEvent = false;
                                    }

                                    //한도금액 체크
                                    SchCls = oDS_PH_PY301B.GetValue("U_SchCls", pVal.Row - 1);
                                    Tuition = Convert.ToDouble(oDS_PH_PY301B.GetValue("U_Tuition", pVal.Row - 1));

                                    //고등학교 이외만 체크
                                    if (Convert.ToInt16(SchCls) > 2)
                                    {
                                        //if (PH_PY301_CheckAmt(Tuition, SchCls) == false)
                                        //{
                                        //    dataHelpClass.MDC_GF_Message("등록금이 한도금액을 초과하였습니다. 확인하십시오.", "E");
                                        //    oDS_PH_PY301B.SetValue("U_Tuition", pVal.Row - 1, Convert.ToString(PreTuition));
                                        //    //이전 데이터로 회귀
                                        //    BubbleEvent = false;
                                        //}
                                    }

                                    //등록금 합계 계산
                                    for (loopCount = 0; loopCount <= oMat1.RowCount - 1; loopCount++)
                                    {
                                        TuiTot += Convert.ToDouble(oDS_PH_PY301B.GetValue("U_Tuition", loopCount));
                                    }
                                    oMat1.LoadFromDataSource();

                                    oDS_PH_PY301A.SetValue("U_TuiTot", 0, Convert.ToString(TuiTot));

                                    FeeTot = Convert.ToDouble(oDS_PH_PY301A.GetValue("U_FeeTot", 0));
                                    Total = FeeTot + TuiTot;

                                    oDS_PH_PY301A.SetValue("U_Total", 0, Convert.ToString(Total));
                                }
                                oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oMat1.AutoResizeColumns();
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "Name" && pVal.CharPressed == 9)
                        {
                            if (string.IsNullOrEmpty(oMat1.Columns.Item("Name").Cells.Item(pVal.Row).Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                    }
                    else if (pVal.ItemUID == "CntcCode" && pVal.CharPressed == 9)
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_KEY_DOWN_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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
                if (pVal.Before_Action == true)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Mat01":
                            if (pVal.Row > 0)
                            {
                                oLastItemUID = pVal.ItemUID;
                                oLastColUID = pVal.ColUID;
                                oLastColRow = pVal.Row;
                            }
                            break;
                        default:
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = "";
                            oLastColRow = 0;
                            break;
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "SchCls")
                            {
                                oMat1.FlushToDataSource();
                                oMat1.LoadFromDataSource();
                            }
                            oMat1.AutoResizeColumns();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_COMBO_SELECT_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat1);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY301A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY301B);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_UNLOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// FORM_RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    oMat1.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_RESIZE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            if (PSH_Globals.SBO_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1293":
                            break;
                        case "1281":
                            break;
                        case "1282":
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            dataHelpClass.AuthorityCheck(oForm, "CLTCOD", "@PH_PY301A", "DocEntry"); //접속자 권한에 따른 사업장 보기
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY301_FormItemEnabled();
                            PH_PY301_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY301_FormItemEnabled();
                            PH_PY301_AddMatrixRow();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            PH_PY301_FormItemEnabled();
                            PH_PY301_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY301_FormItemEnabled();
                            oMat1.AutoResizeColumns();
                            break;
                        case "1293": //행삭제
                            break;
                        case "1287": //복제
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormMenuEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
    }
}

