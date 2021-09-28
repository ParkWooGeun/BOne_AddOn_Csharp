using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 자재기타출고등록
    /// </summary>
    internal class PS_MM090 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_MM090H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_MM090L; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private string oOutMan;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM090.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_MM090_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_MM090");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_MM090_CreateItems();
                PS_MM090_ComboBox_Setting();
                PS_MM090_EnableMenus();
                PS_MM090_SetDocument(oFormDocEntry);

                oForm.EnableMenu(("1287"), true);// 복제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
                oForm.Items.Item("CardCode").Click();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_MM090_CreateItems()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oDS_PS_MM090H = oForm.DataSources.DBDataSources.Item("@PS_MM090H");
                oDS_PS_MM090L = oForm.DataSources.DBDataSources.Item("@PS_MM090L");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" + dataHelpClass.User_MSTCOD() + "'";
                oRecordSet01.DoQuery(sQry);

                oForm.Items.Item("OutMan").Specific.Value = oRecordSet01.Fields.Item(0).Value;
                oOutMan = oRecordSet01.Fields.Item(0).Value;

                oForm.Items.Item("InDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

                oForm.DataSources.UserDataSources.Add("BOM_CHECK", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("BOM_CHECK").Specific.DataBind.SetBound(true, "", "BOM_CHECK");
                oForm.Items.Item("Opt01").Specific.GroupWith("Opt02");
                oForm.Items.Item("Opt01").Click();
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
        /// Combobox 설정
        /// </summary>
        private void PS_MM090_ComboBox_Setting()
        {
            string sQry;
            string CntcCode;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                sQry =  " SELECT  U_Minor AS [Code],";
                sQry += " U_CdName AS [Name]";
                sQry += " FROM [@PS_SY001L]";
                sQry += " WHERE Code = 'M008'";
                sQry += " AND U_UseYN = 'Y'";

                CntcCode = dataHelpClass.User_MSTCOD();

                if (CntcCode == "0074514" || CntcCode == "0672514" || CntcCode == "1178504" || CntcCode == "2160801")
                {
                    sQry += " UNION ALL";
                    sQry += " SELECT  U_Minor AS [Code],";
                    sQry += "         U_CdName AS [Name]";
                    sQry += " FROM    [@PS_SY001L]";
                    sQry += " WHERE   Code = 'M008'";
                    sQry += "         AND U_UseYN = 'N'";
                }
                oForm.Items.Item("OutCls").Specific.ValidValues.Add("%", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("OutCls").Specific, sQry, "", false, false);

                dataHelpClass.Combo_ValidValues_Insert("PS_MM090", "Title", "", "01", "송장");
                dataHelpClass.Combo_ValidValues_Insert("PS_MM090", "Title", "", "02", "거래명세서");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("Title").Specific, "PS_MM090", "Title", false);
                oForm.Items.Item("Title").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL]  order by BPLId", "1", false, false);
                dataHelpClass.Set_ComboList(oForm.Items.Item("Tonnage").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] Where Code = 'M009' and U_UseYN = 'Y'  order by U_Seq", "", true, true);

                oMat01.Columns.Item("OutGbn").ValidValues.Add("", "");
                oMat01.Columns.Item("OutGbn").ValidValues.Add("01", "기타출고");
                oMat01.Columns.Item("OutGbn").ValidValues.Add("10", "MG스크랩출고");
                oMat01.Columns.Item("OutGbn").ValidValues.Add("20", "MG원재료반납");
                oMat01.Columns.Item("OutGbn").ValidValues.Add("30", "MG LOSS");
                oMat01.Columns.Item("OutGbn").ValidValues.Add("40", "내부거래");
                oMat01.Columns.Item("OutGbn").ValidValues.Add("50", "MG스크랩기타출고");
                
                oMat01.Columns.Item("ChkType").ValidValues.Add("", "");
                oMat01.Columns.Item("ChkType").ValidValues.Add("PS_QM450", "검사여부등록");
                oMat01.Columns.Item("ChkType").ValidValues.Add("PS_QM410", "검사성적서(A)등록");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PS_MM090_CalculateAmount
        /// </summary>
        /// <returns></returns>
        private void PS_MM090_CalculateAmount(string pColUID, int pRowID)
        {
            double l_Qty; //수량
            double l_Price; //단가
            double l_Amount; //금액

            try
            {
                oMat01.FlushToDataSource();
                l_Qty = Convert.ToDouble(!string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(pRowID).Specific.Value));
                l_Price = Convert.ToDouble(oMat01.Columns.Item("Price").Cells.Item(pRowID).Specific.Value);
                l_Amount = Convert.ToDouble(oMat01.Columns.Item("Amt").Cells.Item(pRowID).Specific.Value);

                if (pColUID == "Qty")
                {
                    oDS_PS_MM090L.SetValue("U_Amt", pRowID - 1, Convert.ToString(l_Qty * l_Price));
                }
                else if (pColUID == "Price")
                {
                    oDS_PS_MM090L.SetValue("U_Amt", pRowID - 1, Convert.ToString(l_Qty * l_Price));
                }
                else if (pColUID == "Amt")
                {
                    if (l_Qty > 0)
                    {
                        oDS_PS_MM090L.SetValue("U_Price", pRowID - 1, Convert.ToString(l_Amount / l_Qty));
                    }
                }
                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PS_MM090_CheckChkData
        /// </summary>
        /// <param name="pDocDate"></param>
        /// <param name="pOrdNum"></param>
        /// <param name="pChkType"></param>
        /// <param name="pChkDoc"></param>
        /// <param name="pChkNo"></param>
        /// <param name="pDocEntry"></param>
        /// <returns></returns>
        private string PS_MM090_CheckChkData(string pDocDate, string pOrdNum, string pChkType, string pChkDoc, string pChkNo, string pDocEntry)
        {
            string ReturnValue = string.Empty;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = " EXEC PS_MM090_91 '";
                sQry += pDocDate + "','";
                sQry += pOrdNum + "','";
                sQry += pChkType + "','";
                sQry += pChkDoc + "','";
                sQry += pChkNo + "','";
                sQry += pDocEntry + "'";

                oRecordSet01.DoQuery(sQry);

                ReturnValue = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
            return ReturnValue;
        }

        /// <summary>
        /// 처리가능한 Action인지 검사
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PS_MM090_Validate(string ValidateType)
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (ValidateType == "수정")
                {
                }
                else if (ValidateType == "행삭제")
                {
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("LineNum").Cells.Item(oLastColRow01).Specific.Value))
                        {
                        }
                        else
                        {
                            if (oForm.Items.Item("Canceled").Specific.Value == "Y")
                            {
                                errMessage = "취소된문서는 수정할수 없습니다.";
                                throw new Exception();
                            }
                        }
                    }
                }
                else if (ValidateType == "취소")
                {
                    if (oForm.Items.Item("Canceled").Specific.Value == "Y")
                    {
                        errMessage = "이미 취소된 문서입니다.";
                        throw new Exception();
                    }
                }
                returnValue = true;
            }
            catch (Exception ex)
            {
                if(errMessage != string.Empty)
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }
            return returnValue;
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_MM090_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true,true, false, false, false, false, false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFormDocEntry">DocEntry</param>
        private void PS_MM090_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_MM090_FormItemEnabled();
                    PS_MM090_AddMatrixRow(0, true);
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_MM090_FormItemEnabled()
        {
            string User_BPLId;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                User_BPLId = dataHelpClass.User_BPLID();
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("BPLId").Specific.Select(User_BPLId, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("Title").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                    if (User_BPLId == "2")
                    {
                        oForm.Items.Item("OutCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    else
                    {
                        oForm.Items.Item("OutCls").Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    oForm.Items.Item("OutMan").Specific.Value = oOutMan;
                    oForm.Items.Item("InDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

                    PS_MM090_FormClear();

                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가

                    oForm.Items.Item("CardCode").Click();
                    oForm.Items.Item("Btn1").Enabled = false;
                    oForm.Items.Item("Btn02").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = false; // 문서번호
                    oForm.Items.Item("CardCode").Enabled = true; //업체코드
                    oForm.Items.Item("Mat01").Enabled = true; //메트릭스
                    oForm.Items.Item("InDate").Enabled = true; //작성일
                    oForm.Items.Item("BPLId").Enabled = true; //사업장
                    oForm.Items.Item("OutMan").Enabled = true; //반출자
                    oForm.Items.Item("RecMan").Enabled = true; //인수자
                    oForm.Items.Item("PurPose").Enabled = true; //목적
                    oForm.Items.Item("Destin").Enabled = true; //도착지
                    oForm.Items.Item("TranCard").Enabled = true; //운송편
                    oForm.Items.Item("TranCode").Enabled = true; //차량번호
                    oForm.Items.Item("TranCost").Enabled = true; //운임
                    oForm.Items.Item("Title").Enabled = true; //출력타이틀
                    oForm.Items.Item("Opt01").Enabled = true;
                    oForm.Items.Item("Opt02").Enabled = true;
                    oForm.Items.Item("Opt01").Click();
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    //각 모드에 따른 아이템설정
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("DocEntry").Enabled = true; //문서번호활성화
                    oForm.Items.Item("CardCode").Enabled = true; //업체코드
                    oForm.Items.Item("InDate").Enabled = true; //작성일
                    oForm.Items.Item("OutMan").Enabled = true; //반출자
                    oForm.Items.Item("TranCard").Enabled = true; //운송편
                    oForm.Items.Item("Opt01").Enabled = false;
                    oForm.Items.Item("Opt02").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    if (oDS_PS_MM090H.GetValue("Canceled", 0) == "Y")
                    {
                        oForm.Items.Item("Btn1").Enabled = false;  //출력버튼비활성
                        oForm.Items.Item("Btn02").Enabled = false; //출력버튼비활성
                        oForm.Items.Item("DocEntry").Enabled = false;
                        oForm.Items.Item("CardCode").Enabled = false;
                        oForm.Items.Item("Mat01").Enabled = false;
                        oForm.Items.Item("InDate").Enabled = false;
                        oForm.Items.Item("BPLId").Enabled = false;
                        oForm.Items.Item("OutMan").Enabled = false;
                        oForm.Items.Item("RecMan").Enabled = false;
                        oForm.Items.Item("PurPose").Enabled = false;
                        oForm.Items.Item("Destin").Enabled = false;
                        oForm.Items.Item("TranCard").Enabled = false;
                        oForm.Items.Item("TranCode").Enabled = false;
                        oForm.Items.Item("TranCost").Enabled = false;
                        oForm.Items.Item("Title").Enabled = false;
                        oForm.Items.Item("Opt01").Enabled = false;
                        oForm.Items.Item("Opt02").Enabled = false;
                    }
                    else
                    {
                        oForm.Items.Item("Btn1").Enabled = true;
                        oForm.Items.Item("Btn02").Enabled = true;
                        oForm.Items.Item("DocEntry").Enabled = false; //찾기하고나면 문서비활성화처리
                        oForm.Items.Item("CardCode").Enabled = true;
                        oForm.Items.Item("Mat01").Enabled = true;
                        oForm.Items.Item("InDate").Enabled = true;
                        oForm.Items.Item("BPLId").Enabled = true;
                        oForm.Items.Item("OutMan").Enabled = true;
                        oForm.Items.Item("RecMan").Enabled = true;
                        oForm.Items.Item("PurPose").Enabled = true;
                        oForm.Items.Item("Destin").Enabled = true;
                        oForm.Items.Item("TranCard").Enabled = true;
                        oForm.Items.Item("TranCode").Enabled = true;
                        oForm.Items.Item("TranCost").Enabled = true;
                        oForm.Items.Item("Title").Enabled = true;
                        oForm.Items.Item("Opt01").Enabled = false;
                        oForm.Items.Item("Opt02").Enabled = false;
                    }
                    if (oDS_PS_MM090H.GetValue("U_Opt02", 0).ToString().Trim() == "1")
                    {
                        oMat01.Columns.Item("DocLin").Visible = true;
                    }
                    else if (oDS_PS_MM090H.GetValue("U_Opt02", 0).ToString().Trim() == "2")
                    {
                        oMat01.Columns.Item("DocLin").Visible = false;
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
        /// PS_MM090_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_MM090_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_MM090L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_MM090L.Offset = oRow;
                oDS_PS_MM090L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        /// PS_MM090_MTX01
        /// </summary>
        private void PS_MM090_MTX01()
        {
            string errMessage = string.Empty;
            int i;
            string sQty;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            
            try
            {
                oForm.Freeze(true);
                Param01 = oForm.Items.Item("Param01").Specific.Value.ToString().Trim();
                Param02 = oForm.Items.Item("Param01").Specific.Value.ToString().Trim();
                Param03 = oForm.Items.Item("Param01").Specific.Value.ToString().Trim();
                Param04 = oForm.Items.Item("Param01").Specific.Value.ToString().Trim();

                sQty = "SELECT 10";
                oRecordSet01.DoQuery(sQty);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_MM090L.InsertRecord(i);
                    }
                    oDS_PS_MM090L.Offset = i;
                    oDS_PS_MM090L.SetValue("U_COL01", i, oRecordSet01.Fields.Item(0).Value);
                    oDS_PS_MM090L.SetValue("U_COL02", i, oRecordSet01.Fields.Item(1).Value);
                    oRecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                }
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
                if(ProgressBar01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
                }
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_MM090_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM090'", "");
                if (DocEntry == "0")
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PS_MM090_DataValidCheck()
        {
            bool returnValue = false;
            int i = 0;
            string errMessage = string.Empty;
            string ClickCode = string.Empty;
            string type = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value))
                {
                    errMessage = "고객코드는 필수입니다.";
                    type = "F";
                    ClickCode = "CardCode";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("OutMan").Specific.Value))
                {
                    errMessage = "반출자는 필수입니다.";
                    type = "F";
                    ClickCode = "OutMan";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("InDate").Specific.Value))
                {
                    errMessage = "반출일은 필수입니다.";
                    type = "F";
                    ClickCode = "InDate";
                    throw new Exception();
                }
                else if (oForm.Items.Item("OutCls").Specific.Value == "%")
                {
                    errMessage = "반출구분은 필수입니다.";
                    type = "F";
                    ClickCode = "OutCls";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("PurPose").Specific.Value))
                {
                    errMessage = "목적은 필수입니다.";
                    type = "F";
                    ClickCode = "PurPose";
                    throw new Exception();
                }

                if (oMat01.VisualRowCount == 0)
                {
                    errMessage = "라인이 존재하지 않습니다.";
                    type = "X";
                    throw new Exception();
                }
                else
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemName").Cells.Item(1).Specific.Value))
                    {
                        errMessage = "Matrix값이 한줄이상은 있어야합니다.";
                        type = "X";
                        throw new Exception();
                    }

                }
                string CheckData = null;
                for (i = 1; i <= (oMat01.VisualRowCount - 1); i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemName").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "품명은 필수입니다.";
                        type = "M";
                        ClickCode = "ItemName";
                        throw new Exception();
                    }
                    if (oMat01.Columns.Item("OutGbn").Cells.Item(i).Specific.Value == "50")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("PurPose").Specific.Value.ToString().Trim()))
                        {
                            errMessage = "스크랩기타출고는 목적(사유)를 필수입력해야합니다.";
                            type = "M";
                            ClickCode = "PurPose";
                            throw new Exception();
                        }
                    }
                    
                    if (oMat01.Columns.Item("OutGbn").Cells.Item(i).Specific.Value == "40") //내부거래시 적요란에 사용처입력해야함
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("LineNote").Cells.Item(i).Specific.Value.ToString().Trim()))
                        {
                            errMessage = "내부거래시 적요란에 사용처를 필수입력해야합니다.";
                            type = "M";
                            ClickCode = "LineNote";
                            throw new Exception();
                        }
                    }

                    if (oForm.Items.Item("BPLId").Specific.Value != "2")//기계사업부는 라인의 출고구분 필수 미적용
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("OutGbn").Cells.Item(i).Specific.Value))
                        {
                            errMessage = "출고구분은 필수 입니다.";
                            type = "M";
                            ClickCode = "OutGbn";
                            throw new Exception();
                        }
                    }

                    //기계사업부의 납품등록 시 필수 정보 체크(작번, 검사분류, 검사문서번호) 2016.06.17 송명규
                    //납품(제품), 납품(분할납품), 납품(미발주)
                    if (oForm.Items.Item("BPLId").Specific.Value == "2" && (oForm.Items.Item("OutCls").Specific.Value == "01" || oForm.Items.Item("OutCls").Specific.Value == "02" || oForm.Items.Item("OutCls").Specific.Value == "06"))
                    {
                        CheckData = PS_MM090_CheckChkData(oForm.Items.Item("InDate").Specific.Value, oMat01.Columns.Item("JakName").Cells.Item(i).Specific.Value, oMat01.Columns.Item("ChkType").Cells.Item(i).Specific.Value, oMat01.Columns.Item("ChkDoc").Cells.Item(i).Specific.Value, oMat01.Columns.Item("ChkNo").Cells.Item(i).Specific.Value, oForm.Items.Item("DocEntry").Specific.Value);
                        
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("JakName").Cells.Item(i).Specific.Value))
                        {
                            errMessage = i + "행 작번을 입력하십시오.";
                            type = "M";
                            ClickCode = "JakName";
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oMat01.Columns.Item("ChkType").Cells.Item(i).Specific.Value))
                        {
                            errMessage = i + "행 검사분류를 선택하십시오.";
                            type = "M";
                            ClickCode = "ChkType";
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oMat01.Columns.Item("ChkDoc").Cells.Item(i).Specific.Value))
                        {
                            errMessage = i + "행 검사문서번호를 입력하십시오.";
                            type = "M";
                            ClickCode = "ChkDoc";
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oMat01.Columns.Item("ChkNo").Cells.Item(i).Specific.Value))
                        {
                            errMessage = i + "행 성적서번호를 입력하십시오.";
                            type = "M";
                            ClickCode = "ChkNo";
                            throw new Exception();
                        }
                        else if (CheckData != "True")
                        {
                            errMessage = i + "행 작번[" + oMat01.Columns.Item("JakName").Cells.Item(i).Specific.Value + "]의 자료는 " + CheckData;
                            type = "X";
                            throw new Exception();
                        }
                    }
                }

                oMat01.FlushToDataSource();
                oDS_PS_MM090L.RemoveRecord(oDS_PS_MM090L.Size - 1);
                oMat01.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_MM090_FormClear();
                }
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {

                    if (type == "F")
                    {
                        oForm.Items.Item(ClickCode).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                    }
                    else if (type == "M")
                    {
                        oMat01.Columns.Item(ClickCode).Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                    }
                    else
                    {
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            return returnValue;
        }

        /// <summary>
        /// Report_Export
        /// </summary>
        [STAThread]
        private void PS_SMM090_Print_Report01()
        {
            string WinTitle;
            string ReportName;
            string DocEntry;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

                WinTitle = "[PS_MM090] 기타자재출고증출력";
                ReportName = "PS_MM090_01.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

                // Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", DocEntry));

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PS_SMM090_Print_Report02
        /// </summary>
        [STAThread]
        private void PS_SMM090_Print_Report02()
        {
            string WinTitle;
            string ReportName;
            string DocEntry;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                WinTitle = "[PS_MM090] 거래명세서";

                if (oForm.Items.Item("BPLId").Specific.Value == "2")
                {
                    ReportName = "PS_MM090_03.rpt";
                }
                else
                {
                    ReportName = "PS_MM090_02.rpt";
                }

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

                // Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", DocEntry));

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_MM090_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            oOutMan = oForm.Items.Item("OutMan").Specific.Value.ToString().Trim();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {

                            if (PS_MM090_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            oOutMan = oForm.Items.Item("OutMan").Specific.Value.ToString().Trim();
                        }
                    }
                    else if (pVal.ItemUID == "Btn1")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PS_SMM090_Print_Report01);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start(); 
                    }
                    else if (pVal.ItemUID == "Btn02")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PS_SMM090_Print_Report02);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                    else if (pVal.ItemUID == "Opt01")
                    {
                        oMat01.Columns.Item("DocLin").Visible = false;
                        oMat01.AutoResizeColumns();
                    }
                    else if (pVal.ItemUID == "Opt02")
                    {
                        oMat01.Columns.Item("DocLin").Visible = true;
                        oMat01.AutoResizeColumns();
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
                                PS_MM090_FormItemEnabled();
                                PS_MM090_AddMatrixRow(0, true);
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_MM090_FormItemEnabled();
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
        /// KEY_DOWN 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "ItemName")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemName").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    if ((oForm.Items.Item("BPLId").Specific.Selected.Value == "3" || oForm.Items.Item("BPLId").Specific.Selected.Value == "5") && oForm.Items.Item("BOM_CHECK").Specific.Checked == true)
                                    {
                                        PS_SM030 PS_SM030 = new PS_SM030();
                                        PS_SM030.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row, "");
                                        BubbleEvent = false;
                                    }
                                }
                            }
                            else if (pVal.ColUID == "DocLin")
                            {
                                dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "DocLin");
                            }
                        }
                    }
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", ""); //거래처코드
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "OutMan", ""); //반출자조회
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "TranCard", ""); //차량조회
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "ItemCode");
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            string sQry;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.ItemUID == "OutCls")
                {
                    if (oForm.Items.Item("OutCls").Specific.Value == "09")
                    {
                        errMessage = "반출구분이 'MG스크랩'일때는 Line 출고구분행을 'MG스크랩출고'로 등록해주세요.";
                        throw new Exception();
                    }
                }
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            oMat01.FlushToDataSource();
                        }
                        else
                        {
                            if (pVal.ItemUID == "Tonnage")
                            {
                                if (oForm.Items.Item("BPLId").Specific.Value == "1")
                                {
                                    sQry = " Select U_RelCd From [@PS_SY001L] WHERE Code = 'M009' and U_Minor = '" + oForm.Items.Item("Tonnage").Specific.Selected.Value + "'";
                                    oRecordSet01.DoQuery(sQry);
                                    oDS_PS_MM090H.SetValue("U_TranCost", 0, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                                    oDS_PS_MM090H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                                }
                            }
                            else
                            {
                                oDS_PS_MM090H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                            }
                        }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false); //메트릭스 한줄선택시 반전시켜주는 구문
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        oForm.EnableMenu("1281", true); //찾기하고 다시 찾기아이콘활성화처리
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "ItemName")
                            {
                                oDS_PS_MM090L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);

                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_MM090L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_MM090_AddMatrixRow(pVal.Row, false);
                                }
                            }
                            else if (pVal.ColUID == "JakName")
                            {
                                oDS_PS_MM090L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);

                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_MM090L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_MM090_AddMatrixRow(pVal.Row, false);
                                }

                                sQry = " SELECT   FrgnName AS [ItemName],";
                                sQry += " U_Size AS [Size],";
                                sQry += " InvntryUom AS [Unit],";
                                sQry += " 0 AS [Qty]";
                                sQry += " FROM    OITM";
                                sQry += " WHERE   ItemCode = '" + oMat01.Columns.Item("JakName").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";

                                oRecordSet01.DoQuery(sQry);
                                oMat01.Columns.Item("ItemName").Cells.Item(pVal.Row).Specific.Value = oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim();
                                oMat01.Columns.Item("Size").Cells.Item(pVal.Row).Specific.Value = oRecordSet01.Fields.Item("Size").Value.ToString().Trim();
                                oMat01.Columns.Item("Unit").Cells.Item(pVal.Row).Specific.Value = oRecordSet01.Fields.Item("Unit").Value.ToString().Trim();

                            }
                            else if (pVal.ColUID == "DocLin")
                            {
                                oDS_PS_MM090L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);

                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_MM090L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_MM090_AddMatrixRow(pVal.Row, false);
                                }

                                sQry = " EXEC [PS_MM090_90] '" + oMat01.Columns.Item("DocLin").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
                                oRecordSet01.DoQuery(sQry);

                                oDS_PS_MM090L.SetValue("U_JakName", pVal.Row - 1, oRecordSet01.Fields.Item("ItemCode").Value.ToString().Trim());
                                oDS_PS_MM090L.SetValue("U_ItemName", pVal.Row - 1, oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim());
                                oDS_PS_MM090L.SetValue("U_Size", pVal.Row - 1, oRecordSet01.Fields.Item("OutSize").Value.ToString().Trim());
                                oDS_PS_MM090L.SetValue("U_Unit", pVal.Row - 1, oRecordSet01.Fields.Item("OutUnit").Value.ToString().Trim());
                                oDS_PS_MM090L.SetValue("U_Qty", pVal.Row - 1, oRecordSet01.Fields.Item("Qty").Value.ToString().Trim());
                                oDS_PS_MM090L.SetValue("U_Weight", pVal.Row - 1, oRecordSet01.Fields.Item("Qty").Value.ToString().Trim());
                            }
                            else
                            {
                                oDS_PS_MM090L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_MM090H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "CardCode")
                            {
                                oDS_PS_MM090H.SetValue("U_CardName", 0, dataHelpClass.Get_ReData("CardName", "CardCode", "[OCRD]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", ""));
                            }
                        }
                        oMat01.LoadFromDataSource();
                        oMat01.AutoResizeColumns();
                        oForm.Update();
                        if (pVal.ItemUID == "Mat01")
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01")
                    {

                        if (pVal.ColUID == "Qty")
                        {
                            PS_MM090_CalculateAmount(pVal.ColUID, pVal.Row);
                        }
                        else if (pVal.ColUID == "Price")
                        {
                            PS_MM090_CalculateAmount(pVal.ColUID, pVal.Row);
                        }
                        else if (pVal.ColUID == "Amt")
                        {
                            PS_MM090_CalculateAmount(pVal.ColUID, pVal.Row);
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
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
                    PS_MM090_FormItemEnabled();
                    PS_MM090_AddMatrixRow(oMat01.VisualRowCount, false);
                    oMat01.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM090H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM090L);

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PS_MM090_FormItemEnabled();
                    PS_MM090_AddMatrixRow(oMat01.VisualRowCount, false);
                    oMat01.AutoResizeColumns();
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
        /// EVENT_ROW_DELETE
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                int i = 0;
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)//Matrix 행삭제전 행삭제가능여부검사타기
                    {
                        if (PS_MM090_Validate("행삭제") == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }

                        oMat01.FlushToDataSource();

                        oDS_PS_MM090L.RemoveRecord(oDS_PS_MM090L.Size - 1);
                        oMat01.LoadFromDataSource();

                        if (oMat01.RowCount == 0)
                        {
                            PS_MM090_AddMatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_MM090L.GetValue("U_ItemName", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_MM090_AddMatrixRow(oMat01.RowCount, false);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (PS_MM090_Validate("취소") == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            else
                            {
                                dataHelpClass.MDC_GF_Message("현재 모드에서는 취소할수 없습니다.", "W");
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
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
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            PS_MM090_FormItemEnabled();
                            break;
                        case "1282": //추가
                            PS_MM090_FormItemEnabled();
                            PS_MM090_AddMatrixRow(0, true);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PS_MM090_FormItemEnabled();
                            break;
                        case "1287":
                            oDS_PS_MM090H.SetValue("DocEntry", 0, "");
                            oForm.Items.Item("InDate").Specific.Value = "";
                            for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                oMat01.FlushToDataSource();
                                oDS_PS_MM090L.SetValue("DocEntry", i, "");
                                oMat01.LoadFromDataSource();
                            }
                            break;
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
