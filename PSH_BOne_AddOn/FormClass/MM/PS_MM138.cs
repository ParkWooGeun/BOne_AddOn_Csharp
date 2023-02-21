using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 포장 외주입고등록
    /// </summary>
    internal class PS_MM138 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_MM138H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_MM138L; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        private string oBPLId;
        private string oDocdate;
        private string oCardCode;
        private string oCardName;
        private int oDocEntryNext;
        private string oCheck;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM138.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_MM138_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_MM138");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_MM138_CreateItems();
                PS_MM138_ComboBox_Setting();
                PS_MM138_Initialization();
                PS_MM138_FormClear();
                PS_MM138_FormItemEnabled();

                oForm.EnableMenu("1283", false); // 삭제
                oForm.EnableMenu("1285", false); // 복원
                oForm.EnableMenu("1287", true); // 복제
                oForm.EnableMenu("1286", false); // 닫기
                oForm.EnableMenu("1284", true); // 취소
                oForm.EnableMenu("1293", false); // 행삭제
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_MM138_CreateItems()
        {
            try
            {
                oDS_PS_MM138H = oForm.DataSources.DBDataSources.Item("@PS_MM138H");
                oDS_PS_MM138L = oForm.DataSources.DBDataSources.Item("@PS_MM138L");

                oMat01 = oForm.Items.Item("Mat01").Specific;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_MM138_ComboBox_Setting()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
                oRecordSet01.DoQuery(sQry);
                while (!(oRecordSet01.EoF))
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oForm.Items.Item("IssueYN").Specific.ValidValues.Add("Y", "원재료불출완료");
                oForm.Items.Item("IssueYN").Specific.ValidValues.Add("N", "원재료불출대상");
                oForm.Items.Item("IssueYN").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Initialization
        /// </summary>
        private void PS_MM138_Initialization()
        {
            try
            {
                if (string.IsNullOrEmpty(oBPLId))
                {
                    oBPLId = "3";
                }
                if (string.IsNullOrEmpty(oCardCode))
                {
                    oDocdate = DateTime.Now.ToString("yyyyMMdd").Trim();
                }
                oDS_PS_MM138H.SetValue("U_BPLId", 0, oBPLId);
                oDS_PS_MM138H.SetValue("U_DocDate", 0, oDocdate);
                oDS_PS_MM138H.SetValue("U_CardCode", 0, oCardCode);
                oDS_PS_MM138H.SetValue("U_CardName", 0, oCardName);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// HeaderSpaceLineDel
        /// </summary>
        /// <returns></returns>
        private bool PS_MM138_HeaderSpaceLineDel()
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_MM138H.GetValue("U_BPLId", 0)))
                {
                    errMessage = "사업장은 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM138H.GetValue("U_DocDate", 0)))
                {
                    errMessage = "전기일자는 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                // 마감일자 Check
                else if (dataHelpClass.Check_Finish_Status(oDS_PS_MM138H.GetValue("U_BPLId", 0).ToString().Trim(), oDS_PS_MM138H.GetValue("U_DocDate", 0).ToString().Trim().Substring(0, 6)) == false)
                {
                    errMessage = "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM138H.GetValue("U_CardCode", 0)))
                {
                    errMessage = "외주거래처는 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM138H.GetValue("U_CpCode", 0)))
                {
                    errMessage = "공정코드는 필수입력사항입니다. 확인하세요.";
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
        private bool PS_MM138_MatrixSpaceLineDel()
        {
            bool returnValue = false;
            int i;
            string errMessage = string.Empty;

            try
            {
                if (oMat01.VisualRowCount == 0)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }
                else if (oMat01.VisualRowCount == 1)
                {
                    if (string.IsNullOrEmpty(oDS_PS_MM138L.GetValue("U_ItemCode", 0)))
                    {
                        errMessage = "첫라인에 반출문서-행 번호가 없습니다. 확인하세요.";
                        throw new Exception();
                    }
                }
                for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                {
                    if (string.IsNullOrEmpty(oDS_PS_MM138L.GetValue("U_ItemCode", i).ToString().Trim()))
                    {
                        errMessage = "반출문서-행 필수사항입니다. 확인하세요.";
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
        private void PS_MM138_Delete_EmptyRow()
        {
            int i;

            try
            {
                oMat01.FlushToDataSource();
                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oDS_PS_MM138L.GetValue("U_ItemCode", i).ToString().Trim()))
                    {
                        oDS_PS_MM138L.RemoveRecord(i);
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
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_MM138_FormItemEnabled()
        {
            string DocEntry;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("IssueYN").Enabled = true;
                    oForm.Items.Item("ItemCode").Enabled = true;
                    oMat01.Columns.Item("OutQty").Editable = true;
                    oMat01.Columns.Item("OutWt").Editable = true;
                    oForm.Items.Item("IssueYN").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("ItemCode").Enabled = true;
                    oMat01.Columns.Item("OutQty").Editable = false;
                    oMat01.Columns.Item("OutWt").Editable = false;
                    oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("OKYNC").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                    sQry = "Select Count(*) From OCRD Where CardCode = '" + oDS_PS_MM138H.GetValue("U_CardCode", 0).ToString().Trim() + "'";
                    sQry = "select Count(*) From [@PS_MM095H] a Inner Join [@PS_MM095L] b On a.DocEntry = b.DocEntry";
                    sQry += " Where Isnull(a.U_OIGEDoc,'') <> '' and Isnull(a.U_OIGNDoc,'') = '' ";
                    sQry += " And b.U_PP040Doc = '" + DocEntry + "'";

                    oRecordSet01.DoQuery(sQry);
                    if (oRecordSet01.Fields.Item(0).Value > 0)
                    {
                        oForm.Items.Item("IssueYN").Enabled = false;
                    }
                    else
                    {
                        oForm.Items.Item("IssueYN").Enabled = true;
                    }
                    oForm.Items.Item("BPLId").Enabled = false;
                    oForm.Items.Item("CardCode").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = false; // 날짜 수정 못하도록 수정 20180911 황영수
                    oForm.Items.Item("Qty").Enabled = false;
                    oForm.Items.Item("ItemCode").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oMat01.Columns.Item("OutQty").Editable = true;
                    oMat01.Columns.Item("OutWt").Editable = true;
                    oForm.EnableMenu("1284", true);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_MM138_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_MM138_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                if (RowIserted == false)
                {
                    oDS_PS_MM138L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_MM138L.Offset = oRow;
                oDS_PS_MM138L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        /// DocEntry 초기화
        /// </summary>
        private void PS_MM138_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM138'", "");
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_MM138_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            int i;
            string sQry;
            int sRow;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                sRow = oRow;

                switch (oUID)
                {
                    case "CardCode":
                        sQry = "Select CardName From OCRD Where CardCode = '" + oDS_PS_MM138H.GetValue("U_CardCode", 0).ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oDS_PS_MM138H.SetValue("U_CardName", 0, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                        break;

                    case "ItmGrpCd":
                        sQry = "SELECT U_CdName FROM [@PS_SY001L] WHERE Code ='M007' and U_Minor = '" + oForm.Items.Item("ItmGrpCd").Specific.Value.ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oDS_PS_MM138H.SetValue("U_ItmGrpNm", 0, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                        break;

                    case "CpCode":
                        sQry = "Select U_CpName From [@PS_PP001L] Where U_CpCode = '" + oDS_PS_MM138H.GetValue("U_CpCode", 0).ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oDS_PS_MM138H.SetValue("U_CpName", 0, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                        break;

                    case "ItemCode":
                        oMat01.Clear();
                        oForm.Items.Item("Qty").Specific.Value = 0;
                        sQry = "Select ItemName, FrgnName, U_Size From OITM Where ItemCode = '" + oDS_PS_MM138H.GetValue("U_ItemCode", 0).ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oDS_PS_MM138H.SetValue("U_ItemName", 0, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                        sQry = "EXEC [PS_MM138_01] '" + oDS_PS_MM138H.GetValue("U_BPLId", 0).ToString().Trim() + "', '" + oDS_PS_MM138H.GetValue("U_ItemCode", 0).ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oMat01.FlushToDataSource();
                        
                        i = 0;
                        while (!(oRecordSet01.EoF))
                        {
                            oDS_PS_MM138L.InsertRecord(i);
                            oDS_PS_MM138L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                            oDS_PS_MM138L.SetValue("U_ItemCode", i, oRecordSet01.Fields.Item("ItemCode").Value.ToString().Trim());
                            oDS_PS_MM138L.SetValue("U_ItemName", i, oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim());
                            oDS_PS_MM138L.SetValue("U_Size", i, oRecordSet01.Fields.Item("Size").Value.ToString().Trim());
                            oDS_PS_MM138L.SetValue("U_OutItmCd", i, oRecordSet01.Fields.Item("MItemCod").Value.ToString().Trim());
                            oDS_PS_MM138L.SetValue("U_OutItmNm", i, oRecordSet01.Fields.Item("MItemNam").Value.ToString().Trim());
                            oDS_PS_MM138L.SetValue("U_StdQty", i, oRecordSet01.Fields.Item("StdQty").Value.ToString().Trim());
                            oDS_PS_MM138L.SetValue("U_BOMQty", i, oRecordSet01.Fields.Item("BOMQty").Value.ToString().Trim());
                            oDS_PS_MM138L.SetValue("U_BOMWt", i, oRecordSet01.Fields.Item("BOMWt").Value.ToString().Trim());
                            oDS_PS_MM138L.SetValue("U_InWhCd", i, oRecordSet01.Fields.Item("InWhCd").Value.ToString().Trim());
                            oDS_PS_MM138L.SetValue("U_InWhNm", i, oRecordSet01.Fields.Item("InWhNm").Value.ToString().Trim());
                            i += 1;
                            oRecordSet01.MoveNext();
                        }
                        oMat01.LoadFromDataSource();
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// PS_MM138_Add_InventoryGenEntry
        /// </summary>
        /// <returns></returns>
        private bool PS_MM138_Add_InventoryGenEntry()
        {
            bool returnValue = false;
            int RetVal;
            int errDiCode = 0;
            int ResultDocNum;
            string errCode = string.Empty;
            string errDiMsg = string.Empty;
            SAPbobsCOM.Documents oDIObject = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);

            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                PSH_Globals.oCompany.StartTransaction();
                oMat01.FlushToDataSource();

                oDIObject.DocDate = DateTime.ParseExact(oForm.Items.Item("DocDate").Specific.Value, "yyyyMMdd", null);
                oDIObject.UserFields.Fields.Item("U_OrdTyp").Value = "30";
                oDIObject.UserFields.Fields.Item("U_CardCode").Value = oForm.Items.Item("CardCode").Specific.Value;
                oDIObject.UserFields.Fields.Item("U_CardName").Value = oForm.Items.Item("CardName").Specific.Value;
                oDIObject.Lines.ItemCode = oForm.Items.Item("ItemCode").Specific.Value;
                oDIObject.Lines.WarehouseCode = "10" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                oDIObject.Lines.Quantity = Convert.ToDouble(oForm.Items.Item("Qty").Specific.Value);

                oDIObject.Comments = "포장 외주 입고(" + oDS_PS_MM138H.GetValue("DocEntry", 0).ToString().Trim() + ")_PS_MM138";
               
                RetVal = oDIObject.Add();

                if (RetVal != 0)
                {
                    PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                    errCode = "1";
                    throw new Exception();
                }
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    ResultDocNum = Convert.ToInt32(PSH_Globals.oCompany.GetNewObjectKey());
                    oForm.Items.Item("OIGNNo").Specific.Value = ResultDocNum;
                    oDS_PS_MM138H.SetValue("U_OIGNNo", 0, Convert.ToString(ResultDocNum));
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("DI실행 중 오류 발생 : [" + errDiCode + "]" + errDiMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObject);
            }
            return returnValue;
        }

        /// <summary>
        ///  PS_MM138_Add_InventoryGenExit
        /// </summary>
        /// <returns></returns>
        private bool PS_MM138_Add_InventoryGenExit()
        {
            bool returnValue = false;
            int RetVal;
            int errDiCode = 0;
            int ResultDocNum;
            string errCode = string.Empty;
            string errDiMsg = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Documents oDIObject = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);

            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                PSH_Globals.oCompany.StartTransaction();
                oMat01.FlushToDataSource();

                oDIObject.DocDate = DateTime.ParseExact(oForm.Items.Item("DocDate").Specific.Value, "yyyyMMdd", null);
                oDIObject.UserFields.Fields.Item("U_CancDoc").Value = oForm.Items.Item("OIGNNo").Specific.Value.ToString().Trim();
                oDIObject.UserFields.Fields.Item("U_CardCode").Value = oForm.Items.Item("CardCode").Specific.Value;
                oDIObject.UserFields.Fields.Item("U_CardName").Value = oForm.Items.Item("CardName").Specific.Value;
                oDIObject.UserFields.Fields.Item("U_OrdTyp").Value = "30";

                oDIObject.Lines.ItemCode = oForm.Items.Item("ItemCode").Specific.Value;
                oDIObject.Lines.WarehouseCode = "10" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                oDIObject.Lines.Quantity = Convert.ToDouble(oForm.Items.Item("Qty").Specific.Value);

                oDIObject.Comments = "포장 외주 입고 취소(" + oDS_PS_MM138H.GetValue("DocEntry", 0).ToString().Trim() + ")_PS_MM138";

                RetVal = oDIObject.Add();

                if (RetVal != 0)
                {
                    PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
                    errCode = "1";
                    throw new Exception();
                }

                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    ResultDocNum = Convert.ToInt32(PSH_Globals.oCompany.GetNewObjectKey());
                    oForm.Items.Item("OIGENo").Specific.Value = ResultDocNum;
                    oDS_PS_MM138H.SetValue("U_OIGENo", 0, Convert.ToString(ResultDocNum));
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("DI실행 중 오류 발생 : [" + errDiCode + "]" + errDiMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObject);
            }
            return returnValue;
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

                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                //    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            oBPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                            oDocdate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
                            oCardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
                            oCardName = oForm.Items.Item("CardName").Specific.Value.ToString().Trim();

                            if (PS_MM138_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_MM138_MatrixSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            //문서만 생성할때 막음 시작 주석
                            if (PS_MM138_Add_InventoryGenEntry() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            //문서만 생성할때 막음 끝 주석

                            PS_MM138_Delete_EmptyRow();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_MM138_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            //입고문서가 생성안되었을때 DI처리
                            if (string.IsNullOrEmpty(oForm.Items.Item("OIGNNo").Specific.Value.ToString().Trim()))
                            {
                                if (PS_MM138_Add_InventoryGenEntry() == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
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
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PSH_Globals.SBO_Application.ActivateMenuItem("1282");
                        }
                    }
                    else if (pVal.ItemUID == "Btn01")
                    {
                        PSH_Globals.SBO_Application.MessageBox("해당 레포트는 준비되어있지 않습니다.");
                        BubbleEvent = false;
                        return;
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
        /// KEY_DOWN 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "CardCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "OutDoc")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "ItmGrpCd")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "ItemCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "CpCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CpCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "OtDocLin")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                    BubbleEvent = false;
                                }
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
                    if (pVal.ItemUID == "IssueYN")
                    {
                        if (oForm.Items.Item("IssueYN").Specific.Value == "Y")
                        {
                            if (oMat01.VisualRowCount < 1)
                            {
                                PS_MM138_AddMatrixRow(0, true);
                                oMat01.Columns.Item("ItemCode").Cells.Item(1).Specific.Value = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
                                oMat01.Columns.Item("ItemName").Cells.Item(1).Specific.Value = oForm.Items.Item("ItemName").Specific.Value.ToString().Trim();
                            }
                        }
                        else
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                PS_MM138_FlushToItemValue("ItemCode",0,"");
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "CardCode")
                        {
                            PS_MM138_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "OutDoc")
                        {
                            PS_MM138_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "ItmGrpCd")
                        {
                            PS_MM138_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "ItemCode")
                        {
                            PS_MM138_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "CpCode")
                        {
                            PS_MM138_FlushToItemValue(pVal.ItemUID, 0, "");

                        }
                        else if (pVal.ItemUID == "Qty")
                        {
                            if (oMat01.VisualRowCount > 0)
                            {
                                if ((oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() == "3" || oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() == "5") && !string.IsNullOrEmpty(oForm.Items.Item("ItmGrpCd").Specific.Value.ToString().Trim()))
                                {
                                    oMat01.FlushToDataSource();
                                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                    {
                                        if (!string.IsNullOrEmpty(oDS_PS_MM138L.GetValue("U_ItemCode", i).ToString().Trim()))
                                        {
                                            oDS_PS_MM138L.SetValue("U_Qty", i, oForm.Items.Item("Qty").Specific.Value);
                                            if (Convert.ToDouble(oDS_PS_MM138L.GetValue("U_Qty", i)) == Convert.ToDouble(oDS_PS_MM138L.GetValue("U_StdQty", i)))
                                            {
                                                oDS_PS_MM138L.SetValue("U_OutQty", i, oDS_PS_MM138L.GetValue("U_BOMQty", i));
                                                oDS_PS_MM138L.SetValue("U_OutWt", i, oDS_PS_MM138L.GetValue("U_BOMWt", i));
                                            }
                                            else
                                            {
                                                if (codeHelpClass.Left(oDS_PS_MM138H.GetValue("U_ItmGrpCd", 0).ToString().Trim(), 1) == "B")
                                                {
                                                    oDS_PS_MM138L.SetValue("U_OutQty", i, Convert.ToString(Convert.ToDouble(oDS_PS_MM138L.GetValue("U_BOMQty", i)) * (Convert.ToDouble(oDS_PS_MM138L.GetValue("U_Qty", i)) / Convert.ToDouble(oDS_PS_MM138L.GetValue("U_StdQty", i))) / 1000));
                                                    oDS_PS_MM138L.SetValue("U_OutWt", i, Convert.ToString(Convert.ToDouble(oDS_PS_MM138L.GetValue("U_BOMWt", i)) * (Convert.ToDouble(oDS_PS_MM138L.GetValue("U_Qty", i)) / Convert.ToDouble(oDS_PS_MM138L.GetValue("U_StdQty", i))) / 1000));
                                                }
                                                else
                                                {
                                                    oDS_PS_MM138L.SetValue("U_OutQty", i, Convert.ToString(Convert.ToDouble(oDS_PS_MM138L.GetValue("U_BOMQty", i)) * (Convert.ToDouble(oDS_PS_MM138L.GetValue("U_Qty", i)) / Convert.ToDouble(oDS_PS_MM138L.GetValue("U_StdQty", i)))));
                                                    oDS_PS_MM138L.SetValue("U_OutWt", i, Convert.ToString(Convert.ToDouble(oDS_PS_MM138L.GetValue("U_BOMWt", i)) * (Convert.ToDouble(oDS_PS_MM138L.GetValue("U_Qty", i)) / Convert.ToDouble(oDS_PS_MM138L.GetValue("U_StdQty", i)))));
                                                }
                                            }
                                        }
                                    }
                                    oMat01.LoadFromDataSource();
                                    oMat01.Columns.Item("OutQty").Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                else if (oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() == "3" && string.IsNullOrEmpty(oForm.Items.Item("ItmGrpCd").Specific.Value.ToString().Trim()))
                                {
                                    oForm.Items.Item("Qty").Specific.Value = 0;
                                    PSH_Globals.SBO_Application.MessageBox("외주관리그룹을 입력하시기 바랍니다.");
                                }
                            }
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "OtDocLin")
                            {
                                PS_MM138_FlushToItemValue(pVal.ItemUID, pVal.Row,  pVal.ColUID);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM138H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM138L);

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
            string sQry;
            int DocEntry;
            int DocEntryMax;
            int DocEntryNext = 0;
            int OutDoc;
            string BPLID;
            string OIGNNo;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                OIGNNo = oForm.Items.Item("OIGNNo").Specific.Value.ToString().Trim();

                                if (!string.IsNullOrEmpty(OIGNNo.ToString().Trim()))
                                {
                                    BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
                                    OutDoc = Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim());

                                    sQry = "Select Count(*) from [@PS_MM095H] a Inner Join [@PS_MM095L] b On a.DocEntry = b.DocEntry and a.Canceled = 'N' ";
                                    sQry += " Where a.U_BPLId = '" + BPLID + "'";
                                    sQry += " and b.U_WorkGbn = '30' ";
                                    sQry += " and b.U_PP040Doc = '" + OutDoc + "'";

                                    oRecordSet01.DoQuery(sQry);

                                    if (Convert.ToDouble(oRecordSet01.Fields.Item(0).Value) <= 0)
                                    {
                                        if (PSH_Globals.SBO_Application.MessageBox("정말로 취소하시겠습니까?", 1, "예", "아니오") != 1)
                                        {
                                            BubbleEvent = false;
                                            return;
                                        }
                                        if (PS_MM138_Add_InventoryGenExit() == false)
                                        {
                                            BubbleEvent = false;
                                            return;
                                        }
                                    }
                                    else
                                    {
                                        PSH_Globals.SBO_Application.MessageBox("원재료 불출이 되어 취소할 수 없습니다.");
                                        BubbleEvent = false;
                                        return;
                                    }
                                }
                            }
                            else
                            {
                                PSH_Globals.SBO_Application.MessageBox("신규모드일때는 취소할 수 없습니다..");
                                BubbleEvent = false;
                                return;
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
                            if (string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim()))
                            {
                                DocEntry = 0;
                            }
                            else
                            {
                                DocEntry = Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim());
                            }
                            sQry = "Select Max(DocEntry) From [@PS_MM138H]";
                            oRecordSet01.DoQuery(sQry);
                            DocEntryMax = Convert.ToInt32(oRecordSet01.Fields.Item(0).Value.ToString().Trim());

                            if (pVal.MenuUID == "1288")
                            {
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("1290");
                                    BubbleEvent = false;
                                    return;
                                }
                                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                                {
                                    if(string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.Value))
                                    {
                                        PSH_Globals.SBO_Application.ActivateMenuItem("1290");
                                        BubbleEvent = false;
                                        return;
                                    }
                                }
                            One_More_Check_1288:
                                DocEntryNext += 1;
                                if (DocEntryNext > DocEntryMax)
                                {
                                    DocEntry = 0;
                                    goto One_More_Check_1288;
                                }

                                sQry = "Select U_CardCode From [@PS_MM138H] Where DocEntry = '" + DocEntryNext + "'";
                                oRecordSet01.DoQuery(sQry);

                                if (PSH_Globals.oCompany.UserName.ToString().Trim() != oRecordSet01.Fields.Item(0).Value.ToString().Trim())
                                {
                                    if (codeHelpClass.Left(PSH_Globals.oCompany.UserName.ToString().Trim(), 1) == "6")
                                    {
                                        DocEntry += 1;
                                        goto One_More_Check_1288;
                                    }
                                    else
                                    {
                                        oCheck = "False";
                                        return;
                                    }
                                }
                                else
                                {
                                    oCheck = "True";
                                    oDocEntryNext = DocEntryNext;
                                }
                            }
                            else if (pVal.MenuUID == "1289")
                            {
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("1291");
                                    BubbleEvent = false;
                                    return;
                                }
                                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                                {
                                    if (string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.Value))
                                    {
                                        PSH_Globals.SBO_Application.ActivateMenuItem("1291");
                                        BubbleEvent = false;
                                        return;
                                    }
                                }
                            One_More_Check_1289:
                                DocEntryNext = DocEntry - 1;
                                if (DocEntryNext < 1)
                                {
                                    DocEntry = DocEntryMax + 1;
                                    goto One_More_Check_1289;
                                }
                                sQry = "Select U_CardCode From [@PS_MM138H] Where DocEntry = '" + DocEntryNext + "'";
                                oRecordSet01.DoQuery(sQry);

                                if (PSH_Globals.oCompany.UserName.ToString().Trim() != oRecordSet01.Fields.Item(0).Value.ToString().Trim())
                                {
                                    if (codeHelpClass.Left(PSH_Globals.oCompany.UserName.ToString().Trim(), 1) == "6")
                                    {
                                        DocEntry -= 1;
                                        goto One_More_Check_1289;
                                    }
                                    else
                                    {
                                        oCheck = "False";
                                        return;
                                    }
                                }
                                else
                                {
                                    oCheck = "True";
                                    oDocEntryNext = DocEntryNext;
                                }
                            }
                            else if (pVal.MenuUID == "1290")
                            {
                                DocEntryNext = 0;
                            One_More_Check_1290:
                                DocEntryNext += 1;

                                sQry = "Select U_CardCode From [@PS_MM138H] Where DocEntry = '" + DocEntryNext + "'";
                                oRecordSet01.DoQuery(sQry);

                                if (PSH_Globals.oCompany.UserName.ToString().Trim() != oRecordSet01.Fields.Item(0).Value.ToString().Trim())
                                {
                                    if (codeHelpClass.Left(PSH_Globals.oCompany.UserName.ToString().Trim(), 1) == "6")
                                    {
                                        goto One_More_Check_1290;
                                    }
                                    else
                                    {
                                        oCheck = "False";
                                        return;
                                    }
                                }
                                else
                                {
                                    oCheck = "True";
                                    oDocEntryNext = DocEntryNext;
                                }
                            }
                            else if (pVal.MenuUID == "1291")
                            {
                                DocEntryNext += 1;
                            One_More_Check_1291:
                                DocEntryNext -= 1;

                                sQry = "Select U_CardCode From [@PS_MM138H] Where DocEntry = '" + DocEntryNext + "'";
                                oRecordSet01.DoQuery(sQry);

                                if (PSH_Globals.oCompany.UserName.ToString().Trim() != oRecordSet01.Fields.Item(0).Value.ToString().Trim())
                                {
                                    if (codeHelpClass.Left(PSH_Globals.oCompany.UserName.ToString().Trim(), 1) == "6")
                                    {
                                        goto One_More_Check_1291;
                                    }
                                    else
                                    {
                                        oCheck = "False";
                                        return;
                                    }
                                }
                                else
                                {
                                    oCheck = "True";
                                    oDocEntryNext = DocEntryNext;
                                }
                            }
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
                            oDS_PS_MM138H.SetValue("U_BPLId", 0, "3");
                            PS_MM138_FormItemEnabled();
                            if (codeHelpClass.Left(PSH_Globals.oCompany.UserName.ToString().Trim(), 1) == "6")
                            {
                                oForm.Items.Item("CardCode").Specific.Value = PSH_Globals.oCompany.UserName.ToString().Trim();
                                sQry = "Select CardName From OCRD Where CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
                                oRecordSet01.DoQuery(sQry);
                                oForm.Items.Item("CardName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                            }
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                            break;
                        case "1282": //추가
                            PS_MM138_Initialization();
                            PS_MM138_FormClear();
                            PS_MM138_FormItemEnabled();
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
                            if (oCheck == "True")
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("1281");
                                oForm.Items.Item("DocEntry").Specific.Value = oDocEntryNext;
                                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                                oCheck = "False";
                                oDocEntryNext = 0;
                            }

                            PS_MM138_FormItemEnabled();
                            break;
                        case "1287": //복제
                            PS_MM138_FormClear();
                            oDS_PS_MM138H.SetValue("Status", 0, "O");
                            oDS_PS_MM138H.SetValue("Canceled", 0, "N");
                            PS_MM138_FormItemEnabled();
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
    }
}
