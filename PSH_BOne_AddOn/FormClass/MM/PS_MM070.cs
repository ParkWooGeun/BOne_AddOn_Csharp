using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 검수입고
    /// </summary>
    internal class PS_MM070 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_MM070H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_MM070L; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        private int oErrNum;
        private string oQEYesNo;
        private string oPurchase;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM070.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_MM070_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_MM070");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocNum";

                oForm.Freeze(true);
                PS_MM070_CreateItems();
                PS_MM070_ComboBox_Setting();
                PS_MM070_Initialization();
                PS_MM070_FormClear();
                PS_MM070_AddMatrixRow(0, true);
                PS_MM070_FormItemEnabled();

                oForm.EnableMenu("1283", false); //삭제
                oForm.EnableMenu("1287", false); //복제
                oForm.EnableMenu("1286", false); //닫기
                oForm.EnableMenu("1284", true); //취소
                oForm.EnableMenu("1293", true); //행삭제
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
        private void PS_MM070_CreateItems()
        {
            try
            {
                oDS_PS_MM070H = oForm.DataSources.DBDataSources.Item("@PS_MM070H");
                oDS_PS_MM070L = oForm.DataSources.DBDataSources.Item("@PS_MM070L");

                oMat01 = oForm.Items.Item("Mat01").Specific;

                oDS_PS_MM070H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_MM070_ComboBox_Setting()
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";// 사업장
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                sQry = " SELECT '%' AS [Code],"; // 품목구분
                sQry += " '선택' AS [Name]";
                sQry += " UNION ALL";
                sQry += " SELECT Code, ";
                sQry += " Name ";
                sQry += " FROM [@PSH_ORDTYP] ";
                sQry += " ORDER BY Code";

                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("Purchase").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oMat01.Columns.Item("ItemGpCd").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oForm.Items.Item("Purchase").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

                //품질검수여부
                oForm.Items.Item("QEYesNo").Specific.ValidValues.Add("Y", "Yes");
                oForm.Items.Item("QEYesNo").Specific.ValidValues.Add("N", "No");

                //불량코드
                sQry = "SELECT b.U_MidCode, b.U_MidName From [@PS_PP002H] a Inner Join [@PS_PP002L] b On a.DocEntry = b.DocEntry Where a.U_BigCode = '1' Order by b.U_MidCode";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    oMat01.Columns.Item("BadCode1").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
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
        /// Initialization
        /// </summary>
        private void PS_MM070_Initialization()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue); //아이디별 사업장 세팅
                oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD(); //아이디별 사번 세팅

                //품질검수여부
                sQry = "Select dept From [OHEM] Where U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "' ";
                oRecordSet01.DoQuery(sQry);

                if (oRecordSet01.Fields.Item(0).Value.ToString().Trim() == "4" | oRecordSet01.Fields.Item(0).Value.ToString().Trim() == "12")
                {
                    oForm.Items.Item("QEYesNo").Specific.Select("Y");
                }
                else
                {
                    oForm.Items.Item("QEYesNo").Specific.Select("N");
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
        /// HeaderSpaceLineDel
        /// </summary>
        /// <returns></returns>
        private bool PS_MM070_HeaderSpaceLineDel()
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_MM070H.GetValue("U_BPLId", 0)))
                {
                    errMessage = "사업장은 필수입력 사항입니다. 확인하세요.";
                    throw new Exception();
                }
                // 마감일자 Check
                else if (dataHelpClass.Check_Finish_Status(oDS_PS_MM070H.GetValue("U_BPLId", 0).ToString().Trim(), oDS_PS_MM070H.GetValue("U_DocDate", 0).ToString().Trim().Substring(0, 6)) == false)
                {
                    errMessage = "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM070H.GetValue("U_CardCode", 0)))
                {
                    errMessage = "거래처코드는 필수입력 사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_MM070H.GetValue("U_CntcCode", 0)))
                {
                    errMessage = "담당자코드는 필수입력 사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (oDS_PS_MM070H.GetValue("U_Purchase", 0) == "%")
                {
                    errMessage = "구매구분은 필수입력 사항입니다. 확인하세요.";
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
        private bool PS_MM070_MatrixSpaceLineDel()
        {
            bool returnValue = false;
            int i;
            string sQry;
            string errMessage = string.Empty;
            string[] Entry; //가입고문서번호를 저장할 배열변수
            string BaseEntry;
            string BaseLine;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oMat01.FlushToDataSource();
                if (oMat01.VisualRowCount == 0)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }
                for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                {
                    if (string.IsNullOrEmpty(oDS_PS_MM070L.GetValue("U_GADocLin", i)))
                    {
                        errMessage = i + 1 + "번 라인의 가입고문서 - 행이 없습니다.확인하세요.";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oDS_PS_MM070L.GetValue("U_ItemCode", i)))
                    {
                        errMessage = i + 1 + "번 라인의 품목코드가 없습니다. 확인하세요.";
                        throw new Exception();
                    }
                    else if (oDS_PS_MM070H.GetValue("U_Purchase", 0).ToString().Trim() != "30" && oDS_PS_MM070H.GetValue("U_Purchase", 0).ToString().Trim() != "40" && oDS_PS_MM070H.GetValue("U_Purchase", 0).ToString().Trim() != "60" && string.IsNullOrEmpty(oDS_PS_MM070L.GetValue("U_WhsCode", i).ToString().Trim()))
                    {
                        errMessage = i + 1 + "번 라인의 창고코드가 없습니다. 확인하세요.";
                        throw new Exception();
                    }
                    sQry = "Exec [PS_MM070_03] '" + oDS_PS_MM070L.GetValue("U_GADocLin", i) + "','" + oDS_PS_MM070L.GetValue("U_ItemCode", i) + "'";
                    oRecordSet.DoQuery(sQry);

                    if(oRecordSet.Fields.Item(0).Value - Convert.ToInt32(oDS_PS_MM070L.GetValue("U_Qty", i)) < 0)
                    {   
                        errMessage = "가입고수량보다 검수입고를 초과하실 수 없습니다. " + Environment.NewLine + "해당 품목 : " + oDS_PS_MM070L.GetValue("U_ItemCode", i) + Environment.NewLine + "검수입고가능잔량 : " + oRecordSet.Fields.Item(0).Value ;
                        throw new Exception();
                    }
                    if (oDS_PS_MM070L.GetValue("U_InQty", i) == oDS_PS_MM070L.GetValue("U_BadQty", i))
                    {
                    }
                    else
                    {
                        if (Convert.ToDouble(oDS_PS_MM070L.GetValue("U_Weight", i)) == 0)
                        {
                            errMessage = i + 1 + "번 라인의 이론중량이 0입니다. 확인하세요.";
                            throw new Exception();
                        }
                        if (Convert.ToDouble(oDS_PS_MM070L.GetValue("U_RealWt", i)) == 0)
                        {
                            errMessage = i + 1 + "번 라인의 실중량이 0입니다. 확인하세요.";
                            throw new Exception();
                        }
                        if (Convert.ToDouble(oDS_PS_MM070L.GetValue("U_Price", i)) == 0)
                        {
                            errMessage = i + 1 + "번 라인의 단가는 0보다 커야 합니다. 확인하세요.";
                            throw new Exception();
                        }
                        if (Convert.ToDouble(oDS_PS_MM070L.GetValue("U_LinTotal", i)) == 0)
                        {
                            errMessage = i + 1 + "번 라인의 금액은 0보다 커야 합니다. 확인하세요.";
                            throw new Exception();
                        }
                    }

                    if (Convert.ToDouble(oDS_PS_MM070L.GetValue("U_BadQty", i)) > 0 && !string.IsNullOrEmpty(oDS_PS_MM070L.GetValue("U_BadQty", i)))
                    {
                        if (string.IsNullOrEmpty(oDS_PS_MM070L.GetValue("U_BadCode1", i)) || string.IsNullOrEmpty(oDS_PS_MM070L.GetValue("U_BadCode2", i)))
                        {
                            errMessage = i + 1 + "번 라인에 불량수량이 있으니 불량중분류와 불량상세를 입력해야 합니다. 확인하세요.";
                            throw new Exception();
                        }
                    }
                    Entry = oMat01.Columns.Item("GADocLin").Cells.Item(i + 1).Specific.Value.Split('-');
                    BaseEntry = Entry[0];
                    BaseLine = Entry[1];

                    if (PS_MM070_CheckDate(BaseEntry) == false) //가입고와 일자 체크
                    {
                        errMessage = i + 1 + "행 [" + oMat01.Columns.Item("ItemCode").Cells.Item(i + 1).Specific.Value + "]의 검수입고일은 가입고일과 같거나 늦어야합니다. 확인하십시오." + Environment.NewLine + "해당 검수입고는 전체가 등록되지 않습니다.";
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
        private void PS_MM070_Delete_EmptyRow()
        {
            int i;
            string errMessage = string.Empty;

            try
            {
                oMat01.FlushToDataSource();
                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oDS_PS_MM070L.GetValue("U_ItemCode", i).ToString().Trim()))
                    {
                        oDS_PS_MM070L.RemoveRecord(i); // Mat01에 마지막라인(빈라인) 삭제
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
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_MM070_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = false;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("QEYesNo").Enabled = true;
                    oForm.Items.Item("Purchase").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("DueDate").Enabled = true;
                    oMat01.Columns.Item("GADocLin").Editable = true;
                    oMat01.Columns.Item("ItemCode").Editable = true;
                    oMat01.Columns.Item("ItemGpCd").Editable = true;
                    oMat01.Columns.Item("BatchYN").Editable = true;
                    oMat01.Columns.Item("BatchNum").Editable = true;
                    oMat01.Columns.Item("InQty").Editable = true;
                    oMat01.Columns.Item("Qty").Editable = true;
                    oMat01.Columns.Item("BadQty").Editable = false;
                    oMat01.Columns.Item("Weight").Editable = true;
                    oMat01.Columns.Item("RealWt").Editable = true;
                    oMat01.Columns.Item("BadCode1").Editable = true;
                    oMat01.Columns.Item("BadCode2").Editable = true;
                    oMat01.Columns.Item("UnWeight").Editable = true;
                    oMat01.Columns.Item("LinTotal").Editable = true;
                    oMat01.Columns.Item("WhsCode").Editable = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = true;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("QEYesNo").Enabled = true;
                    oForm.Items.Item("Purchase").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("DueDate").Enabled = true;
                    oMat01.Columns.Item("GADocLin").Editable = false;
                    oMat01.Columns.Item("ItemCode").Editable = false;
                    oMat01.Columns.Item("ItemGpCd").Editable = false;
                    oMat01.Columns.Item("BatchYN").Editable = false;
                    oMat01.Columns.Item("BatchNum").Editable = false;
                    oMat01.Columns.Item("InQty").Editable = false;
                    oMat01.Columns.Item("Qty").Editable = false;
                    oMat01.Columns.Item("BadQty").Editable = false;
                    oMat01.Columns.Item("Weight").Editable = false;
                    oMat01.Columns.Item("RealWt").Editable = false;
                    oMat01.Columns.Item("BadCode1").Editable = false;
                    oMat01.Columns.Item("BadCode2").Editable = false;
                    oMat01.Columns.Item("UnWeight").Editable = false;
                    oMat01.Columns.Item("LinTotal").Editable = false;
                    oMat01.Columns.Item("WhsCode").Editable = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("DocNum").Enabled = false;
                    oForm.Items.Item("CardCode").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = false;
                    oForm.Items.Item("QEYesNo").Enabled = false;
                    oForm.Items.Item("Purchase").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = false;
                    oForm.Items.Item("DueDate").Enabled = false;
                    oMat01.Columns.Item("GADocLin").Editable = false;
                    oMat01.Columns.Item("ItemCode").Editable = false;
                    oMat01.Columns.Item("ItemGpCd").Editable = false;
                    oMat01.Columns.Item("BatchYN").Editable = false;
                    oMat01.Columns.Item("BatchNum").Editable = false;
                    oMat01.Columns.Item("InQty").Editable = false;
                    oMat01.Columns.Item("Qty").Editable = false;
                    oMat01.Columns.Item("BadQty").Editable = false;
                    oMat01.Columns.Item("Weight").Editable = false;
                    oMat01.Columns.Item("RealWt").Editable = false;
                    oMat01.Columns.Item("BadCode1").Editable = false;
                    oMat01.Columns.Item("BadCode2").Editable = false;
                    oMat01.Columns.Item("UnWeight").Editable = false;
                    oMat01.Columns.Item("LinTotal").Editable = false;
                    oMat01.Columns.Item("WhsCode").Editable = false;

                    if (string.IsNullOrEmpty(oMat01.Columns.Item("BatchNum").Cells.Item(1).Specific.Value))
                    {
                        oForm.Items.Item("Btn_prt").Enabled = false;
                    }
                    else
                    {
                        oForm.Items.Item("Btn_prt").Enabled = true;
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
        /// PS_MM070_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_MM070_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_MM070L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_MM070L.Offset = oRow;
                oDS_PS_MM070L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        private void PS_MM070_FormClear()
        {
            string DocNum;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocNum = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM070'", "");
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
        private void PS_MM070_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            int Qty;
            int GAQty;
            string sQry;
            string ItemCode;
            double BadQty;
            double GARealWt;
            double GAUnWt;
            double Price;
            double RealWt;
            double Weight;
            double Calculate_Weight;
            double BadWt;
            string GADocLin;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                switch (oUID)
                {
                    case "CardCode":
                        sQry = "Select CardName From OCRD Where CardCode = '" + oDS_PS_MM070H.GetValue("U_CardCode", 0).ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oDS_PS_MM070H.SetValue("U_CardName", 0, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                        break;

                    case "CntcCode":
                        sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" + oDS_PS_MM070H.GetValue("U_CntcCode", 0).ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        oDS_PS_MM070H.SetValue("U_CntcName", 0, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                        break;

                    case "Mat01":
                        if (oCol == "GADocLin")
                        {
                            if (oRow == oMat01.RowCount || oMat01.VisualRowCount == 0 && !string.IsNullOrEmpty(oMat01.Columns.Item("GADocLin").Cells.Item(oRow).Specific.Value.ToString().Trim()))
                            {
                                oMat01.FlushToDataSource();
                                PS_MM070_AddMatrixRow(oMat01.RowCount, false);
                                oMat01.Columns.Item("GADocLin").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                            GADocLin = oDS_PS_MM070L.GetValue("U_GADocLin", oRow - 1).ToString().Trim();

                            sQry = " Select  a.U_CardCode, ";
                            sQry += " b.U_ItemCode,";
                            sQry += " b.U_ItemName,";
                            sQry += " b.U_FrgnName,";
                            sQry += " b.U_ItemGpCd,";
                            sQry += " b.U_QTy,";
                            sQry += " b.U_Weight,";
                            sQry += " b.U_RealWt,";
                            sQry += " b.U_UnWeight,";
                            sQry += " b.U_Price,";
                            sQry += " b.U_LinTotal,";
                            sQry += " b.U_WhsCode,";
                            sQry += " b.U_WhsName,";
                            sQry += " b.U_BDocNum,";
                            sQry += " b.U_BLineNum,";
                            sQry += " a.U_DocDate,";
                            sQry += " b.U_OutSize,";
                            sQry += " b.U_OutUnit,";
                            sQry += " b.U_Auto,";
                            sQry += " b.U_DocCur,"; //통화
                            sQry += " b.U_DocRate,"; //환율
                            sQry += " b.U_FCPrice,"; //외화환산단가
                            sQry += " b.U_FCAmount,"; //외화환산금액
                            sQry += " case ";
                            sQry += "     when c.U_ItmBsort = '317' then c.ManBtchNum";
                            sQry += "     else ''";
                            sQry += " end as ManBtchNum,";
                            sQry += " case ";
                            sQry += "     when c.ManBtchNum ='Y' and c.U_ItmBsort = '317' then right(b.u_itemcode,5)";
                            sQry += "     else ''";
                            sQry += " end";
                            sQry += " +";
                            sQry += " case ";
                            sQry += "     when c.ManBtchNum ='Y' and c.U_ItmBsort = '317' then";
                            sQry += "         convert(varchar(6),getdate(),12)";
                            sQry += "         +";
                            sQry += "         (";
                            sQry += "             select      convert(varchar(1),case when count(right(d.U_BatchNum,1)) ='0' then '1' else right(max(d.U_BatchNum),1) + 1 end) ";
                            sQry += "             From        [@PS_MM070H] c";
                            sQry += "                         Inner Join";
                            sQry += "                         [@PS_MM070L] d";
                            sQry += "                             On c.DocEntry = d.DocEntry";
                            sQry += "                             and c.canceled ='N' ";
                            sQry += "             where       c.U_DocDate = convert(varchar(8),getdate(),112)";
                            sQry += "                         and d.u_itemcode = b.U_ItemCode";
                            sQry += "         )  else ''";
                            sQry += " end as BatchNum ";
                            sQry += " From    [@PS_MM050H] a";
                            sQry += "         Inner Join";
                            sQry += "         [@PS_MM050L] b";
                            sQry += "             On a.DocEntry = b.DocEntry";
                            sQry += "         left Join";
                            sQry += "         OITM c ";
                            sQry += "             On b.U_ItemCode = C.ItemCode ";
                            sQry += " Where   a.DocNum = Left('" + GADocLin + "', CharIndex('-', '" + GADocLin + "') - 1)";
                            sQry += "         And b.LineId = Right('" + GADocLin + "', Len('" + GADocLin + "') - CharIndex('-', '" + GADocLin + "'))"; //U_LineNum을 LineId로 수정(2012.07.24 송명규)

                            oRecordSet01.DoQuery(sQry);

                            if (string.IsNullOrEmpty(oDS_PS_MM070H.GetValue("U_CardCode", 0).ToString().Trim()))
                            {
                                oForm.Items.Item("CardCode").Specific.Value = oRecordSet01.Fields.Item("U_CardCode").Value.ToString().Trim();
                            }
                            else
                            {
                                if (oDS_PS_MM070H.GetValue("U_CardCode", 0).ToString().Trim() != oRecordSet01.Fields.Item("U_CardCode").Value.ToString().Trim())
                                {
                                    errMessage = "다른 거래처의 데이터를 한문서에서 처리할 수 없습니다. 확인하세요.";
                                    throw new Exception();
                                }
                            }

                            oMat01.FlushToDataSource();

                            if (oRecordSet01.RecordCount == 0)
                            {
                                //매트릭스에 데이터를 직접 바인딩하면 이벤트가 실행되기 때문에 DataSource로 바인딩하는 방식으로 수정(2011.11.22 송명규)
                                oDS_PS_MM070L.SetValue("U_ItemCode", oRow - 1, "");
                                oDS_PS_MM070L.SetValue("U_ItemName", oRow - 1, "");
                                oDS_PS_MM070L.SetValue("U_OutSize", oRow - 1, "");
                                oDS_PS_MM070L.SetValue("U_OutUnit", oRow - 1, "");
                                oDS_PS_MM070L.SetValue("U_InQty", oRow - 1, "0");
                                oDS_PS_MM070L.SetValue("U_Qty", oRow - 1, "0");
                                oDS_PS_MM070L.SetValue("U_Weight", oRow - 1, "0");
                                oDS_PS_MM070L.SetValue("U_RealWt", oRow - 1, "0");
                                oDS_PS_MM070L.SetValue("U_Auto", oRow - 1, "N");
                                oDS_PS_MM070L.SetValue("U_OPORNum", oRow - 1, "");
                                oDS_PS_MM070L.SetValue("U_POR1Num", oRow - 1, "");

                                dataHelpClass.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.:", "W");
                                oMat01.LoadFromDataSource();
                                return;
                            }

                            oForm.Items.Item("DueDate").Specific.Value = oRecordSet01.Fields.Item("U_DocDate").Value.ToString("yyyyMMdd").Trim();

                            //매트릭스에 데이터를 직접 바인딩하면 이벤트가 실행되기 때문에 DataSource로 바인딩하는 방식으로 수정(2011.11.22 송명규)
                            oDS_PS_MM070L.SetValue("U_ItemCode", oRow - 1, oRecordSet01.Fields.Item("U_ItemCode").Value.ToString().Trim()); //품목코드
                            oDS_PS_MM070L.SetValue("U_ItemName", oRow - 1, oRecordSet01.Fields.Item("U_ItemName").Value.ToString().Trim()); //품목이름
                            oDS_PS_MM070L.SetValue("U_OutSize", oRow - 1, oRecordSet01.Fields.Item("U_OutSize").Value.ToString().Trim()); //규격
                            oDS_PS_MM070L.SetValue("U_OutUnit", oRow - 1, oRecordSet01.Fields.Item("U_OutUnit").Value.ToString().Trim()); //단위
                            oDS_PS_MM070L.SetValue("U_ItemGpCd", oRow - 1, oRecordSet01.Fields.Item("U_ItemGpCd").Value.ToString().Trim()); //구매구분
                            oDS_PS_MM070L.SetValue("U_GAQty", oRow - 1, oRecordSet01.Fields.Item("U_Qty").Value.ToString().Trim()); //가입고수량
                            oDS_PS_MM070L.SetValue("U_GARealWt", oRow - 1, oRecordSet01.Fields.Item("U_RealWt").Value.ToString().Trim()); //가입고실중량
                            oDS_PS_MM070L.SetValue("U_InQty", oRow - 1, oRecordSet01.Fields.Item("U_Qty").Value.ToString().Trim()); //입고수량
                            oDS_PS_MM070L.SetValue("U_Qty", oRow - 1, oRecordSet01.Fields.Item("U_Qty").Value.ToString().Trim()); //검수수량
                            oDS_PS_MM070L.SetValue("U_Weight", oRow - 1, oRecordSet01.Fields.Item("U_Weight").Value.ToString().Trim()); //이론중량
                            oDS_PS_MM070L.SetValue("U_RealWt", oRow - 1, oRecordSet01.Fields.Item("U_RealWt").Value.ToString().Trim()); //실중량
                            oDS_PS_MM070L.SetValue("U_BadQty", oRow - 1, "0"); //불량수량
                            oDS_PS_MM070L.SetValue("U_BatchYN", oRow - 1, oRecordSet01.Fields.Item("ManBtchNum").Value.ToString().Trim()); //불량수량
                            oDS_PS_MM070L.SetValue("U_BatchNum", oRow - 1, oRecordSet01.Fields.Item("BatchNum").Value.ToString().Trim()); //불량수량
                            oDS_PS_MM070L.SetValue("U_UnWeight", oRow - 1, oRecordSet01.Fields.Item("U_UnWeight").Value.ToString().Trim()); //단중
                            oDS_PS_MM070L.SetValue("U_Price", oRow - 1, oRecordSet01.Fields.Item("U_Price").Value.ToString().Trim()); //단가
                            oDS_PS_MM070L.SetValue("U_LinTotal", oRow - 1, oRecordSet01.Fields.Item("U_LinTotal").Value.ToString().Trim()); //금액
                            oDS_PS_MM070L.SetValue("U_WhsCode", oRow - 1, oRecordSet01.Fields.Item("U_WhsCode").Value.ToString().Trim()); //창고코드
                            oDS_PS_MM070L.SetValue("U_WhsName", oRow - 1, oRecordSet01.Fields.Item("U_WhsName").Value.ToString().Trim()); //창고명(2011.11.22 송명규 추가)
                            oDS_PS_MM070L.SetValue("U_Auto", oRow - 1, oRecordSet01.Fields.Item("U_Auto").Value.ToString().Trim()); //자동불출여부
                            oDS_PS_MM070L.SetValue("U_OPORNum", oRow - 1, oRecordSet01.Fields.Item("U_BDocNum").Value.ToString().Trim()); //구매오더번호
                            oDS_PS_MM070L.SetValue("U_POR1Num", oRow - 1, oRecordSet01.Fields.Item("U_BLineNum").Value.ToString().Trim()); //오더행
                            oDS_PS_MM070L.SetValue("U_DocCur", oRow - 1, oRecordSet01.Fields.Item("U_DocCur").Value.ToString().Trim()); //통화
                            oDS_PS_MM070L.SetValue("U_DocRate", oRow - 1, oRecordSet01.Fields.Item("U_DocRate").Value.ToString().Trim()); //환율
                            oDS_PS_MM070L.SetValue("U_FCPrice", oRow - 1, oRecordSet01.Fields.Item("U_FCPrice").Value.ToString().Trim()); //외화환산단가
                            oDS_PS_MM070L.SetValue("U_FCAmount", oRow - 1, oRecordSet01.Fields.Item("U_FCAmount").Value.ToString().Trim()); //외화환산금액

                            oMat01.LoadFromDataSource();
                            oMat01.AutoResizeColumns();
                        }
                        else if (oCol == "ItemCode")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                if (oRow == oMat01.RowCount || oMat01.VisualRowCount == 0 && !string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim()))
                                {
                                    oMat01.FlushToDataSource();
                                    PS_MM070_AddMatrixRow(oMat01.RowCount, false);
                                }
                            }
                            sQry = "Select ItemName, FrgnName From OITM Where ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);

                            //매트릭스에 데이터를 직접 바인딩하면 이벤트가 실행되기 때문에 DataSource로 바인딩하는 방식으로 수정(2011.11.22 송명규)
                            oMat01.FlushToDataSource();
                            oDS_PS_MM070L.SetValue("U_ItemName", oRow - 1, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                            oMat01.LoadFromDataSource();
                            oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else if (oCol == "WhsCode")
                        {
                            sQry = "Select WhsName From [OWHS] Where WhsCode = '" + oMat01.Columns.Item("WhsCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);

                            //매트릭스에 데이터를 직접 바인딩하면 이벤트가 실행되기 때문에 DataSource로 바인딩하는 방식으로 수정(2011.11.22 송명규)
                            oMat01.FlushToDataSource();

                            oDS_PS_MM070L.SetValue("U_WhsName", oRow - 1, oRecordSet01.Fields.Item(0).Value.ToString().Trim());

                            oMat01.LoadFromDataSource();
                            oMat01.Columns.Item("WhsCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else if (oCol == "InQty")
                        {
                            oMat01.FlushToDataSource();
                            if (!string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(oRow).Specific.Value.ToString().Trim()))
                            {
                                BadQty = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_InQty", oRow - 1)) - Convert.ToDouble(oDS_PS_MM070L.GetValue("U_Qty", oRow - 1));
                                oDS_PS_MM070L.SetValue("U_BadQty", oRow - 1, Convert.ToString(BadQty));
                            }
                            else
                            {
                                oDS_PS_MM070L.SetValue("U_BadQty", oRow - 1, "0");
                            }
                            oMat01.LoadFromDataSource();
                            oMat01.Columns.Item("InQty").Cells.Item(oRow).Click();
                        }
                        else if (oCol == "Qty")
                        {
                            oMat01.FlushToDataSource();

                            //중량계산
                            //서브작번이 추가된 작번을 넘기면 안되므로 Main 작번만 넘김(2011.10.25 송명규 수정)
                            ItemCode = codeHelpClass.Left(oDS_PS_MM070L.GetValue("U_ItemCode", oRow - 1).ToString().Trim(), 11);
                            Qty = Convert.ToInt32(oDS_PS_MM070L.GetValue("U_Qty", oRow - 1));

                            Calculate_Weight = dataHelpClass.Calculate_Weight(ItemCode, Qty, oForm.Items.Item("BPLId").Specific.Value.ToString().Trim());

                            oDS_PS_MM070L.SetValue("U_Weight", oRow - 1, Convert.ToString(Calculate_Weight));

                            //불량수량계산
                            if (!string.IsNullOrEmpty(oDS_PS_MM070L.GetValue("U_InQty", oRow - 1)))
                            {
                                BadQty = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_InQty", oRow - 1)) - Convert.ToDouble(oDS_PS_MM070L.GetValue("U_Qty", oRow - 1));
                                oDS_PS_MM070L.SetValue("U_BadQty", oRow - 1, Convert.ToString(BadQty));
                                //불량중량 = 가입고중량 - 이론중량 (2011.11.03 송명규 추가)
                                BadWt = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_GARealWt", oRow - 1)) - Convert.ToDouble(oDS_PS_MM070L.GetValue("U_Weight", oRow - 1));
                                oDS_PS_MM070L.SetValue("U_BadWt", oRow - 1, Convert.ToString(BadWt));
                            }
                            else
                            {
                                oDS_PS_MM070L.SetValue("U_BadQty", oRow - 1, "0");
                                oDS_PS_MM070L.SetValue("U_BadWt", oRow - 1, "0");
                            }

                            //실중량계산
                            GAQty = Convert.ToInt32(oDS_PS_MM070L.GetValue("U_GAQty", oRow - 1));
                            GARealWt = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_GARealWt", oRow - 1));

                            if (!string.IsNullOrEmpty(oDS_PS_MM070L.GetValue("U_GAQty", oRow - 1)) && GAQty != 0)
                            {
                                GAUnWt = System.Math.Round(GARealWt / GAQty, 2);
                            }
                            else
                            {
                                GAUnWt = 0;
                            }
                            RealWt = GAUnWt * Convert.ToDouble(oDS_PS_MM070L.GetValue("U_Qty", oRow - 1));

                            sQry = "Select ItmsGrpCod From OITM Where ItemCode = '" + ItemCode + "'";
                            oRecordSet01.DoQuery(sQry);
                            if (oRecordSet01.Fields.Item(0).Value.ToString().Trim() == "105")
                            {
                                oDS_PS_MM070L.SetValue("U_RealWt", oRow - 1, Convert.ToString(Calculate_Weight));
                            }
                            else
                            {
                                oDS_PS_MM070L.SetValue("U_RealWt", oRow - 1, Convert.ToString(RealWt));
                            }

                            //금액 계산(2012.05.03 송명규 추가)
                            //단가X이론중량
                            Price = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_Price", oRow - 1));
                            Weight = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_Weight", oRow - 1));
                            oDS_PS_MM070L.SetValue("U_LinTotal", oRow - 1, Convert.ToString(Weight * Price));

                            oMat01.LoadFromDataSource();
                            oMat01.Columns.Item("BadQty").Cells.Item(oRow).Click();
                        }
                        else if (oCol == "Weight")
                        {
                            if (!string.IsNullOrEmpty(oMat01.Columns.Item("Weight").Cells.Item(oRow).Specific.Value.ToString().Trim()))
                            {
                                oMat01.FlushToDataSource();
                                Price = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_Price", oRow - 1));
                                Weight = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_Weight", oRow - 1));
                                oDS_PS_MM070L.SetValue("U_LinTotal", oRow - 1, Convert.ToString(Weight * Price));

                                oMat01.LoadFromDataSource();
                                oMat01.Columns.Item("Weight").Cells.Item(oRow).Click();
                            }
                        }
                        break;
                }
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
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 선행프로세스와 일자 비교
        /// </summary>
        /// <param name="pBaseEntry">기준문서번호</param>
        /// <returns>선행프로세스보다 일자가 같거나 느릴 경우(true), 선행프로세스보다 일자가 빠를 경우(false)</returns>
        private bool PS_MM070_CheckDate(string pBaseEntry)
        {
            bool returnValue = false;
            string query;
            string BaseEntry;
            string BaseLine;
            string DocType;
            string CurDocDate;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                BaseEntry = pBaseEntry;
                BaseLine = "";
                DocType = "PS_MM070";
                CurDocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();

                query = "EXEC PS_Z_CHECK_DATE '";
                query += BaseEntry + "','";
                query += BaseLine + "','";
                query += DocType + "','";
                query += CurDocDate + "'";

                oRecordSet01.DoQuery(query);

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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
            return returnValue;
        }

        /// <summary>
        /// PS_MM070_Add_oPurchaseDeliveryNotes
        /// </summary>
        /// <returns></returns>
        private bool PS_MM070_Add_oPurchaseDeliveryNotes()
        {
            bool returnValue = false;
            int i;
            int j;
            int errCode;
            int RetVal;
            int Line_Count;
            string ErrMsg;
            string sQry;
            string DueDate;
            string DocDate;
            string DocEntry;
            string ECVatGroup;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Documents DI_oPurchaseDeliveryNotes = null;
            SAPbobsCOM.Documents DI_oInventoryGenExit = null;

            try
            {
                DocDate = dataHelpClass.ConvertDateType(oDS_PS_MM070H.GetValue("U_DocDate", 0), "-");
                DueDate = dataHelpClass.ConvertDateType(oDS_PS_MM070H.GetValue("U_DueDate", 0), "-");
                if (string.IsNullOrEmpty(oDS_PS_MM070H.GetValue("U_GRDocNum", 0).ToString().Trim()))
                {
                    if (PSH_Globals.oCompany.InTransaction == true)
                    {
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    }
                    PSH_Globals.oCompany.StartTransaction();
                    DI_oPurchaseDeliveryNotes = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes); //'입고PO 문서객체
                    oMat01.FlushToDataSource(); 

                    if (string.IsNullOrEmpty(oDS_PS_MM070H.GetValue("U_GRDocNum", 0).ToString().Trim()))
                    {
                        DI_oPurchaseDeliveryNotes.CardCode = oDS_PS_MM070H.GetValue("U_CardCode", 0).ToString().Trim();
                        DI_oPurchaseDeliveryNotes.BPL_IDAssignedToInvoice = Convert.ToInt32(oDS_PS_MM070H.GetValue("U_BPLId", 0));
                        DI_oPurchaseDeliveryNotes.DocDate = Convert.ToDateTime(DocDate);
                        DI_oPurchaseDeliveryNotes.DocDueDate = Convert.ToDateTime(DueDate);
                        DI_oPurchaseDeliveryNotes.Comments = "검수입고(" + oDS_PS_MM070H.GetValue("DocNum", 0).ToString().Trim() + ")입고PO : " + oDS_PS_MM070H.GetValue("U_Comments", 0).ToString().Trim();

                        sQry = "Select ECVatGroup From [OCRD] Where CardCode = '" + oDS_PS_MM070H.GetValue("U_CardCode", 0).ToString().Trim() + "'";
                        oRecordSet01.DoQuery(sQry);
                        ECVatGroup = oRecordSet01.Fields.Item("ECVatGroup").Value;

                        Line_Count = 0;
                        
                        if (oDS_PS_MM070H.GetValue("U_Purchase", 0).ToString().Trim() == "30" || oDS_PS_MM070H.GetValue("U_Purchase", 0).ToString().Trim() == "40" || oDS_PS_MM070H.GetValue("U_Purchase", 0).ToString().Trim() == "60") //가공비, 외주제작, 고정자산 품의 일때는 Service로 생성
                        {
                            DI_oPurchaseDeliveryNotes.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;
                            DI_oPurchaseDeliveryNotes.DocCurrency = oDS_PS_MM070L.GetValue("U_DocCur", 0).ToString().Trim(); //ISO 통화기호 추가(2020.04.06 송명규)
                            DI_oPurchaseDeliveryNotes.DocRate = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_DocRate", 0).ToString().Trim()); //환율 추가(2020.04.06 송명규)
                            for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                            {
                                if (Convert.ToDouble(oDS_PS_MM070L.GetValue("U_Weight", i).ToString().Trim()) <= 0)
                                {
                                    continue;
                                }
                                if (i > 0)
                                {
                                    DI_oPurchaseDeliveryNotes.Lines.Add();
                                }
                                DI_oPurchaseDeliveryNotes.Lines.SetCurrentLine(i);

                                //외화처리 기능 구헌(2020.04.06 송명규)_S
                                if (oDS_PS_MM070L.GetValue("U_DocCur", 0).ToString().Trim() == "KRW")
                                {
                                    DI_oPurchaseDeliveryNotes.Lines.LineTotal = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_LinTotal", i).ToString().Trim());
                                }
                                else
                                {
                                    DI_oPurchaseDeliveryNotes.Lines.LineTotal = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_FCAmount", i).ToString().Trim());
                                }
                                //외화처리 기능 구헌(2020.04.06 송명규)_E

                                DI_oPurchaseDeliveryNotes.Lines.VatGroup = ECVatGroup;

                                //품의구분에 따른 GL계정 분류(2012.07.09 송명규 수정)
                                if (oDS_PS_MM070H.GetValue("U_Purchase", 0).ToString().Trim() == "60")
                                {
                                    DI_oPurchaseDeliveryNotes.Lines.AccountCode = "59109020"; //가공비품의
                                }
                                else if (oDS_PS_MM070H.GetValue("U_Purchase", 0).ToString().Trim() == "30")
                                {
                                    DI_oPurchaseDeliveryNotes.Lines.AccountCode = "55302010"; //외주가공비(제) 계정
                                }
                                else if (oDS_PS_MM070H.GetValue("U_Purchase", 0).ToString().Trim() == "40")
                                {
                                    DI_oPurchaseDeliveryNotes.Lines.AccountCode = "11506100"; //원재료 계정
                                }
                                //품의구분에 따른 GL계정 분류(2012.07.09 송명규 수정)
                                DI_oPurchaseDeliveryNotes.Lines.UserFields.Fields.Item("U_sItemCode").Value = oDS_PS_MM070L.GetValue("U_ItemCode", i).ToString().Trim();
                                DI_oPurchaseDeliveryNotes.Lines.UserFields.Fields.Item("U_sItemName").Value = oDS_PS_MM070L.GetValue("U_ItemName", i).ToString().Trim();
                                DI_oPurchaseDeliveryNotes.Lines.UserFields.Fields.Item("U_sSize").Value = oDS_PS_MM070L.GetValue("U_OutSize", i).ToString().Trim();
                                DI_oPurchaseDeliveryNotes.Lines.UserFields.Fields.Item("U_sUnit").Value = oDS_PS_MM070L.GetValue("U_OutUnit", i).ToString().Trim();
                                DI_oPurchaseDeliveryNotes.Lines.UserFields.Fields.Item("U_sQty").Value = oDS_PS_MM070L.GetValue("U_Qty", i).ToString().Trim();
                                DI_oPurchaseDeliveryNotes.Lines.UserFields.Fields.Item("U_sWeight").Value = oDS_PS_MM070L.GetValue("U_Weight", i).ToString().Trim();

                                DI_oPurchaseDeliveryNotes.Lines.BaseType = 22;
                                DI_oPurchaseDeliveryNotes.Lines.BaseEntry = Convert.ToInt32(oDS_PS_MM070L.GetValue("U_OPORNum", i).ToString().Trim());
                                DI_oPurchaseDeliveryNotes.Lines.BaseLine = Convert.ToInt32(oDS_PS_MM070L.GetValue("U_POR1Num", i).ToString().Trim());

                                sQry = "Select b.U_Payment From OPOR a inner Join POR1 b On a.DocEntry = b.DocEntry Where a.DocEntry = '" + oDS_PS_MM070L.GetValue("U_OPORNum", i).ToString().Trim() + "' and b.LineNum = '" + oDS_PS_MM070L.GetValue("U_POR1Num", i).ToString().Trim() + "'";
                                oRecordSet01.DoQuery(sQry);

                                DI_oPurchaseDeliveryNotes.Lines.UserFields.Fields.Item("U_Payment").Value = (oRecordSet01.Fields.Item("U_Payment").Value);

                                Line_Count += 1;
                            }
                        }
                        else
                        {
                            DI_oPurchaseDeliveryNotes.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                            DI_oPurchaseDeliveryNotes.DocCurrency = oDS_PS_MM070L.GetValue("U_DocCur", 0).ToString().Trim(); //ISO 통화기호 추가(2020.04.06 송명규)
                            DI_oPurchaseDeliveryNotes.DocRate = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_DocRate", 0).ToString().Trim()); //환율 추가(2020.04.06 송명규)

                            for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                            {
                                if (Convert.ToDouble(oDS_PS_MM070L.GetValue("U_RealWt", i).ToString().Trim()) <= 0)
                                {
                                    continue;
                                }
                                if (i > 0)
                                {
                                    DI_oPurchaseDeliveryNotes.Lines.Add();
                                }
                            
                                DI_oPurchaseDeliveryNotes.Lines.SetCurrentLine(i);
                                DI_oPurchaseDeliveryNotes.Lines.ItemCode = oDS_PS_MM070L.GetValue("U_ItemCode", i).ToString().Trim();
                                DI_oPurchaseDeliveryNotes.Lines.Quantity = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_RealWt", i).ToString().Trim());
                                DI_oPurchaseDeliveryNotes.Lines.DiscountPercent = 100 - (Convert.ToDouble(oDS_PS_MM070L.GetValue("U_LinTotal", i).ToString().Trim()) / (Convert.ToDouble(oDS_PS_MM070L.GetValue("U_RealWt", i).ToString().Trim()) * Convert.ToDouble(oDS_PS_MM070L.GetValue("U_Price", i).ToString().Trim())) * 100);
                                
                                //외화처리 기능 구헌(2020.04.06 송명규)_S
                                if (oDS_PS_MM070L.GetValue("U_DocCur", 0).ToString().Trim() == "KRW")
                                {
                                    DI_oPurchaseDeliveryNotes.Lines.UnitPrice = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_Price", i).ToString().Trim());
                                    DI_oPurchaseDeliveryNotes.Lines.LineTotal = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_LinTotal", i).ToString().Trim());
                                }
                                else
                                {
                                    DI_oPurchaseDeliveryNotes.Lines.UnitPrice = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_FCPrice", i).ToString().Trim());
                                    DI_oPurchaseDeliveryNotes.Lines.LineTotal = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_FCAmount", i).ToString().Trim());
                                }
                                //외화처리 기능 구헌(2020.04.06 송명규)_E

                                DI_oPurchaseDeliveryNotes.Lines.WarehouseCode = oDS_PS_MM070L.GetValue("U_WhsCode", i).ToString().Trim();
                                DI_oPurchaseDeliveryNotes.Lines.UserFields.Fields.Item("U_Qty").Value = oDS_PS_MM070L.GetValue("U_Qty", i).ToString().Trim();
                                DI_oPurchaseDeliveryNotes.Lines.UserFields.Fields.Item("U_Weight").Value = oDS_PS_MM070L.GetValue("U_Weight", i).ToString().Trim();
                                DI_oPurchaseDeliveryNotes.Lines.UserFields.Fields.Item("U_sPrice").Value = oDS_PS_MM070L.GetValue("U_Price", i).ToString().Trim();
                                //단가 추가(2011.11.28 송명규 추가)

                                DI_oPurchaseDeliveryNotes.Lines.BaseType = Convert.ToInt32("22");
                                DI_oPurchaseDeliveryNotes.Lines.BaseEntry = Convert.ToInt32(oDS_PS_MM070L.GetValue("U_OPORNum", i).ToString().Trim());
                                DI_oPurchaseDeliveryNotes.Lines.BaseLine = Convert.ToInt32(oDS_PS_MM070L.GetValue("U_POR1Num", i).ToString().Trim());

                                if (oDS_PS_MM070L.GetValue("U_BatchYN", i).ToString().Trim() == "Y")
                                {
                                    DI_oPurchaseDeliveryNotes.Lines.BatchNumbers.BatchNumber = oDS_PS_MM070L.GetValue("U_BatchNum", i).ToString().Trim();
                                    DI_oPurchaseDeliveryNotes.Lines.BatchNumbers.Quantity = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_Weight", i).ToString().Trim());
                                }

                                sQry = "Select b.U_Payment From OPOR a inner Join POR1 b On a.DocEntry = b.DocEntry Where a.DocEntry = '" + oDS_PS_MM070L.GetValue("U_OPORNum", i).ToString().Trim() + "' and b.LineNum = '" + oDS_PS_MM070L.GetValue("U_POR1Num", i).ToString().Trim() + "'";
                                oRecordSet01.DoQuery(sQry);

                                DI_oPurchaseDeliveryNotes.Lines.UserFields.Fields.Item("U_Payment").Value = oRecordSet01.Fields.Item("U_Payment").Value.ToString().Trim();
                                Line_Count += 1;
                            }
                        }
                        
                        //자동불출(출고 DIAPI)
                        if (Line_Count > 0)
                        {
                            j = 0;
                            if (oDS_PS_MM070H.GetValue("U_Purchase", 0).ToString().Trim() == "30" || oDS_PS_MM070H.GetValue("U_Purchase", 0).ToString().Trim() == "40" || oDS_PS_MM070H.GetValue("U_Purchase", 0).ToString().Trim() == "60")
                            {
                            }
                            else
                            {
                                if (string.IsNullOrEmpty(oDS_PS_MM070H.GetValue("U_OIGENum", 0).ToString().Trim()))
                                {
                                    for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                                    {
                                        if (oDS_PS_MM070L.GetValue("U_Auto", i).ToString().Trim() == "Y" && Convert.ToDouble(oDS_PS_MM070L.GetValue("U_RealWt", i).ToString().Trim()) > 0)
                                        {
                                            if (j == 0)
                                            {
                                                DI_oInventoryGenExit = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);

                                                DI_oInventoryGenExit.BPL_IDAssignedToInvoice = Convert.ToInt32(oDS_PS_MM070H.GetValue("U_BPLId", 0).ToString().Trim());
                                                DI_oInventoryGenExit.DocDate = Convert.ToDateTime(DocDate);
                                                DI_oInventoryGenExit.TaxDate = Convert.ToDateTime(DocDate);
                                                DI_oInventoryGenExit.Comments = "검수입고(" + oDS_PS_MM070H.GetValue("DocNum", 0).ToString().Trim() + ")출고 : " + oDS_PS_MM070H.GetValue("U_Comments", 0).ToString().Trim();
                                            }

                                            if (j > 0) 
                                            {
                                                DI_oInventoryGenExit.Lines.Add();
                                            }

                                            sQry = "EXEC [PS_MM070_01] '" + oDS_PS_MM070L.GetValue("U_GADocLin", i).ToString().Trim() + "'";
                                            oRecordSet01.DoQuery(sQry);

                                            DI_oInventoryGenExit.Lines.SetCurrentLine(j);
                                            DI_oInventoryGenExit.Lines.ItemCode = oDS_PS_MM070L.GetValue("U_ItemCode", i).ToString().Trim();
                                            DI_oInventoryGenExit.Lines.WarehouseCode = oDS_PS_MM070L.GetValue("U_WhsCode", i).ToString().Trim();
                                            DI_oInventoryGenExit.Lines.Quantity = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_RealWt", i).ToString().Trim());
                                            DI_oInventoryGenExit.Lines.CostingCode = oRecordSet01.Fields.Item(0).Value.ToString().Trim(); //배부규칙

                                            //금액, 단가 지정하여 입력(2014.01.29 송명규 추가)
                                            DI_oInventoryGenExit.Lines.LineTotal = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_LinTotal", i).ToString().Trim()); //금액
                                            DI_oInventoryGenExit.Lines.Price = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_Price", i)); //단가

                                            j += 1;
                                        }
                                    }
                                }
                            }
                        }

                        if (Line_Count != 0)
                        {
                            RetVal = DI_oPurchaseDeliveryNotes.Add();
                            if (0 != RetVal)
                            {
                                PSH_Globals.oCompany.GetLastError(out errCode, out ErrMsg);
                                errMessage = ErrMsg;
                                throw new Exception();
                            }

                            PSH_Globals.oCompany.GetNewObjectCode(out DocEntry);
                            oDS_PS_MM070H.SetValue("U_GRDocNum", 0, DocEntry);

                            if (DI_oInventoryGenExit == null)
                            {
                            }
                            else
                            {
                                RetVal = DI_oInventoryGenExit.Add();
                                if (0 != RetVal)
                                {
                                    PSH_Globals.oCompany.GetLastError(out errCode, out ErrMsg);
                                    errMessage = ErrMsg;
                                    throw new Exception();
                                }
                                PSH_Globals.oCompany.GetNewObjectCode(out DocEntry);
                                oDS_PS_MM070H.SetValue("U_OIGENum", 0, DocEntry);
                            }
                            PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        }
                    }
                }
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
                if (DI_oInventoryGenExit != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oInventoryGenExit);
                }
                if (DI_oPurchaseDeliveryNotes != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oPurchaseDeliveryNotes);
                }
            }
            return returnValue;
        }

        /// <summary>
        /// PS_MM070_Add_oPurchaseReturns
        /// </summary>
        /// <returns></returns>
        private bool PS_MM070_Add_oPurchaseReturns()
        {
            bool returnValue = false;
            int i = 0;
            int j;
            int errCode;
            int RetVal;
            string ErrMsg;
            string sQry;
            string DueDate;
            string DocDate;
            string DocEntry;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Documents DI_oPurchaseReturns = null; //구매오더 문서객체
            SAPbobsCOM.Documents DI_oInventoryGenEntry = null; //입고 문서객체

            try
            {
                DocDate = dataHelpClass.ConvertDateType(oDS_PS_MM070H.GetValue("U_DocDate", 0), "-");
                DueDate = dataHelpClass.ConvertDateType(oDS_PS_MM070H.GetValue("U_DueDate", 0), "-");

                if (!string.IsNullOrEmpty(oDS_PS_MM070H.GetValue("U_GRDocNum", 0).ToString().Trim()) && string.IsNullOrEmpty(oDS_PS_MM070H.GetValue("U_ORPDNum", 0).ToString().Trim()))
                {
                    if (PSH_Globals.oCompany.InTransaction == true)
                    {
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    }
                    PSH_Globals.oCompany.StartTransaction();
                    DI_oPurchaseReturns = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseReturns);

                    DI_oPurchaseReturns.CardCode = oDS_PS_MM070H.GetValue("U_CardCode", 0).ToString().Trim();
                    DI_oPurchaseReturns.BPL_IDAssignedToInvoice = Convert.ToInt32(oDS_PS_MM070H.GetValue("U_BPLId", 0).ToString().Trim());
                    DI_oPurchaseReturns.DocDate = Convert.ToDateTime(DocDate);
                    DI_oPurchaseReturns.DocDueDate = Convert.ToDateTime(DueDate);
                    DI_oPurchaseReturns.Comments = "검수입고(" + oDS_PS_MM070H.GetValue("DocNum", 0).ToString().Trim() + ")반품 : " + oDS_PS_MM070H.GetValue("U_Comments", 0);
                    DI_oPurchaseReturns.UserFields.Fields.Item("U_CancDoc").Value = oDS_PS_MM070H.GetValue("U_GRDocNum", 0).ToString().Trim();

                    sQry = " Select T1.ItemCode,";
                    sQry += " T1.Dscription,";
                    sQry += " T1.WhsCode,";
                    sQry += " T1.Quantity,";
                    sQry += " T1.Price,";
                    sQry += " T1.DiscPrcnt,";
                    sQry += " T1.LineTotal,";
                    sQry += " T1.U_Weight,";
                    sQry += " T0.ObjType,";
                    sQry += " T1.DocEntry,";
                    sQry += " T1.LineNum,";
                    sQry += " T1.U_sItemCode,";
                    sQry += " T1.U_SItemName,";
                    sQry += " T1.U_sSize,";
                    sQry += " T1.U_sUnit,";
                    sQry += " T1.U_sQty,";
                    sQry += " T1.U_sWeight,";
                    sQry += " T1.VatGroup";
                    sQry += " From        OPDN T0 Inner Join PDN1 T1 On T1.DocEntry = T0.DocEntry ";
                    sQry += " Where       T0.DocNum = '" + oDS_PS_MM070H.GetValue("U_GRDocNum", 0).ToString().Trim() + "'";
                    sQry += " Order by    T1.LineNum";

                    oRecordSet01.DoQuery(sQry);

                    if (oDS_PS_MM070H.GetValue("U_Purchase", i).ToString().Trim() == "30" || oDS_PS_MM070H.GetValue("U_Purchase", i).ToString().Trim() == "40" || oDS_PS_MM070H.GetValue("U_Purchase", i).ToString().Trim() == "60")
                    {
                        DI_oPurchaseReturns.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;

                        while (!oRecordSet01.EoF)
                        {
                            if (i > 0)
                            {
                                DI_oPurchaseReturns.Lines.Add();
                            }
                            DI_oPurchaseReturns.Lines.SetCurrentLine(i);
                            DI_oPurchaseReturns.Lines.ItemDescription = oRecordSet01.Fields.Item("Dscription").Value.ToString().Trim();
                            DI_oPurchaseReturns.Lines.LineTotal = Convert.ToDouble(oRecordSet01.Fields.Item("LineTotal").Value.ToString().Trim());
                            DI_oPurchaseReturns.Lines.VatGroup = oRecordSet01.Fields.Item("VatGroup").Value.ToString().Trim();
                            DI_oPurchaseReturns.Lines.UserFields.Fields.Item("U_sItemCode").Value = oRecordSet01.Fields.Item("U_sItemCode").Value.ToString().Trim();
                            DI_oPurchaseReturns.Lines.UserFields.Fields.Item("U_sItemName").Value = oRecordSet01.Fields.Item("U_sItemName").Value.ToString().Trim();
                            DI_oPurchaseReturns.Lines.UserFields.Fields.Item("U_sSize").Value = oRecordSet01.Fields.Item("U_sSize").Value.ToString().Trim();
                            DI_oPurchaseReturns.Lines.UserFields.Fields.Item("U_sUnit").Value = oRecordSet01.Fields.Item("U_sUnit").Value.ToString().Trim();
                            DI_oPurchaseReturns.Lines.UserFields.Fields.Item("U_sQty").Value = oRecordSet01.Fields.Item("U_sQty").Value.ToString().Trim();
                            DI_oPurchaseReturns.Lines.UserFields.Fields.Item("U_sWeight").Value = oRecordSet01.Fields.Item("U_sWeight").Value.ToString().Trim();
                            DI_oPurchaseReturns.Lines.BaseType = Convert.ToInt32(oRecordSet01.Fields.Item("ObjType").Value.ToString().Trim());
                            DI_oPurchaseReturns.Lines.BaseEntry = Convert.ToInt32(oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());
                            DI_oPurchaseReturns.Lines.BaseLine = Convert.ToInt32(oRecordSet01.Fields.Item("LineNum").Value.ToString().Trim());
                            i += 1;
                            oRecordSet01.MoveNext();
                        }
                    }
                    else
                    {
                        DI_oPurchaseReturns.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                        while (!oRecordSet01.EoF)
                        {
                            if (i > 0)
                            {
                                DI_oPurchaseReturns.Lines.Add();
                            }
                            DI_oPurchaseReturns.Lines.SetCurrentLine(i);
                            DI_oPurchaseReturns.Lines.ItemCode = oRecordSet01.Fields.Item("ItemCode").Value.ToString().Trim();
                            DI_oPurchaseReturns.Lines.Quantity = Convert.ToDouble(oRecordSet01.Fields.Item("Quantity").Value.ToString().Trim());
                            DI_oPurchaseReturns.Lines.Price = Convert.ToDouble(oRecordSet01.Fields.Item("Price").Value.ToString().Trim());
                            DI_oPurchaseReturns.Lines.DiscountPercent = Convert.ToDouble(oRecordSet01.Fields.Item("DiscPrcnt").Value.ToString().Trim());
                            DI_oPurchaseReturns.Lines.LineTotal = Convert.ToDouble(oRecordSet01.Fields.Item("LineTotal").Value.ToString().Trim());
                            DI_oPurchaseReturns.Lines.WarehouseCode = oRecordSet01.Fields.Item("WhsCode").Value.ToString().Trim();
                            DI_oPurchaseReturns.Lines.UserFields.Fields.Item("U_Qty").Value = oDS_PS_MM070L.GetValue("U_Qty", i).ToString().Trim();
                            DI_oPurchaseReturns.Lines.UserFields.Fields.Item("U_Weight").Value = oRecordSet01.Fields.Item("U_Weight").Value.ToString().Trim();

                            if (oDS_PS_MM070L.GetValue("U_BatchYN", i).ToString().Trim() == "Y")
                            {
                                DI_oPurchaseReturns.Lines.BatchNumbers.BatchNumber = oDS_PS_MM070L.GetValue("U_BatchNum", i).ToString().Trim();
                                DI_oPurchaseReturns.Lines.BatchNumbers.Quantity = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_Weight", i).ToString().Trim());
                            }
                            DI_oPurchaseReturns.Lines.BaseType = Convert.ToInt32(oRecordSet01.Fields.Item("ObjType").Value.ToString().Trim());
                            DI_oPurchaseReturns.Lines.BaseEntry = Convert.ToInt32(oRecordSet01.Fields.Item("DocEntry").Value.ToString().Trim());
                            DI_oPurchaseReturns.Lines.BaseLine = Convert.ToInt32(oRecordSet01.Fields.Item("LineNum").Value.ToString().Trim());
                            i += 1;
                            oRecordSet01.MoveNext();
                        }
                    }

                    j = 0;
                    if (oDS_PS_MM070H.GetValue("U_Purchase", 0).ToString().Trim() == "30" || oDS_PS_MM070H.GetValue("U_Purchase", 0).ToString().Trim() == "40" || oDS_PS_MM070H.GetValue("U_Purchase", 0).ToString().Trim() == "60")
                    {
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(oDS_PS_MM070H.GetValue("U_OIGENum", 0).ToString().Trim()) && string.IsNullOrEmpty(oDS_PS_MM070H.GetValue("U_OIGNNum", 0).ToString().Trim()))
                        {
                            for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                            {
                                if (oDS_PS_MM070L.GetValue("U_Auto", i).ToString().Trim() == "Y" && Convert.ToDouble(oDS_PS_MM070L.GetValue("U_RealWt", i).ToString().Trim()) > 0)
                                {
                                    if (j == 0)
                                    {
                                        DI_oInventoryGenEntry = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
                                        DI_oInventoryGenEntry.BPL_IDAssignedToInvoice = Convert.ToInt32(oDS_PS_MM070H.GetValue("U_BPLId", 0).ToString().Trim());
                                        DI_oInventoryGenEntry.DocDate = Convert.ToDateTime(DocDate);
                                        DI_oInventoryGenEntry.TaxDate = Convert.ToDateTime(DocDate);
                                        DI_oInventoryGenEntry.Comments = "검수입고(" + oDS_PS_MM070H.GetValue("DocNum", 0).ToString().Trim() + ")출고 : " + oDS_PS_MM070H.GetValue("U_Comments", 0).ToString().Trim();
                                    }

                                    if (j > 0) 
                                    {
                                        DI_oInventoryGenEntry.Lines.Add();
                                    }
                                        
                                    sQry = "EXEC [PS_MM070_01] '" + oDS_PS_MM070L.GetValue("U_GADocLin", i).ToString().Trim() + "'";
                                    oRecordSet01.DoQuery(sQry);

                                    DI_oInventoryGenEntry.Lines.SetCurrentLine(j);
                                    DI_oInventoryGenEntry.Lines.ItemCode = oDS_PS_MM070L.GetValue("U_ItemCode", i).ToString().Trim();
                                    DI_oInventoryGenEntry.Lines.WarehouseCode = oDS_PS_MM070L.GetValue("U_WhsCode", i).ToString().Trim();
                                    DI_oInventoryGenEntry.Lines.Quantity = Convert.ToDouble(oDS_PS_MM070L.GetValue("U_RealWt", i).ToString().Trim());
                                    DI_oInventoryGenEntry.Lines.CostingCode = oRecordSet01.Fields.Item(0).Value.ToString().Trim();

                                    j += 1;
                                }
                            }
                        }
                    }

                    if (DI_oInventoryGenEntry == null)
                    {
                    }
                    else
                    {
                        RetVal = DI_oInventoryGenEntry.Add();
                        if (0 != RetVal)
                        {
                            PSH_Globals.oCompany.GetLastError(out errCode, out ErrMsg);
                            errMessage = ErrMsg;
                            throw new Exception();
                        }

                        PSH_Globals.oCompany.GetNewObjectCode(out DocEntry);
                        oDS_PS_MM070H.SetValue("U_OIGNNum", 0, DocEntry);
                    }

                    RetVal = DI_oPurchaseReturns.Add();
                    if (0 != RetVal)
                    {
                        PSH_Globals.oCompany.GetLastError(out errCode, out ErrMsg);
                        errMessage = ErrMsg;
                        throw new Exception();
                    }
                    PSH_Globals.oCompany.GetNewObjectCode(out DocEntry);
                    oDS_PS_MM070H.SetValue("U_ORPDNum", 0, DocEntry);

                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                }
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
                if (DI_oPurchaseReturns != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oPurchaseReturns);
                }
                if (DI_oInventoryGenEntry != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oInventoryGenEntry);
                }
            }
            return returnValue;
        }

        /// <summary>
        /// PS_MM070_Add_oPurchaseReturns
        /// </summary>
        /// <returns></returns>
        private bool PS_MM070_Close_oPurchaseDeliveryNotes()
        {
            bool returnValue = false;
            int RetVal; 
            int errCode;
            string ErrMsg;
            string DocEntry;
            string errMessage = string.Empty;
            SAPbobsCOM.Documents DI_oPurchaseDeliveryNotes = null;

            try
            {
                DocEntry = oDS_PS_MM070H.GetValue("U_PODocNum", 0).ToString().Trim();

                if (!string.IsNullOrEmpty(oDS_PS_MM070H.GetValue("U_PODocNum", 0).ToString().Trim()))
                {
                    PSH_Globals.oCompany.StartTransaction();
                    DI_oPurchaseDeliveryNotes = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);

                    if (DI_oPurchaseDeliveryNotes.GetByKey(Convert.ToInt32(DocEntry)) == false)
                    {
                        PSH_Globals.oCompany.GetLastError(out errCode, out ErrMsg);
                        errMessage = ErrMsg;
                        throw new Exception();
                    }
                    RetVal = DI_oPurchaseDeliveryNotes.Close();
                    if (0 != RetVal)
                    {
                        PSH_Globals.oCompany.GetLastError(out errCode, out ErrMsg);
                        throw new Exception();
                    }

                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                }
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
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
                if (DI_oPurchaseDeliveryNotes != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oPurchaseDeliveryNotes);
                }
            }
            return returnValue;
        }

        /// <summary>
        /// 처리가능한 Action인지 검사
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PS_MM070_Validate(string ValidateType)
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
                                errMessage = "취소된 문서는 수정할수 없습니다.";
                                throw new Exception();
                            }
                            else
                            {
                                errMessage = "등록된 문서는 행삭제할 수 없습니다. 취소 후 다시 등록하세요.";
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }
            return returnValue;
        }

        /// <summary>
        /// Report_Export
        /// </summary>
        [STAThread]
        private void PS_MM070_Print_Report01()
        {
            string WinTitle;
            string ReportName;
            string DocEntry;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                WinTitle = "[PS_MM070_01] 라벨 출력";
                ReportName = "PS_MM070_01.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

                dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", DocEntry));// Parameter

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
            int j;
            string batchNum;
            string vReturnValue;

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_MM070_HeaderSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            if (PS_MM070_MatrixSpaceLineDel() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                batchNum = oMat01.Columns.Item("BatchNum").Cells.Item(i + 1).Specific.Value;
                                for (j = 0; j <= oMat01.VisualRowCount - 1; j++)
                                {
                                    if (i != j && batchNum == oMat01.Columns.Item("BatchNum").Cells.Item(j + 1).Specific.Value && !string.IsNullOrEmpty(oMat01.Columns.Item("BatchNum").Cells.Item(j + 1).Specific.Value))
                                    {
                                        PSH_Globals.SBO_Application.MessageBox("중복");
                                        oMat01.Columns.Item("BatchNum").Cells.Item(j + 1).Specific.Value = Convert.ToDouble(batchNum) + 1;
                                    }
                                }
                            }

                            if (DateTime.ParseExact(oForm.Items.Item("DocDate").Specific.Value, "yyyyMMdd", null) < DateTime.ParseExact(oForm.Items.Item("DueDate").Specific.Value, "yyyyMMdd", null))
                            {
                                vReturnValue = Convert.ToString(PSH_Globals.SBO_Application.MessageBox("납품일보다 검수일이 빠릅니다. 계속하겠습니까?", 1, "&확인", "&취소"));
                                if (Convert.ToDouble(vReturnValue) == 1)
                                {
                                    //검수입고 문서만 등록 시 주석처리_S
                                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                    {
                                        if (PS_MM070_Add_oPurchaseDeliveryNotes() == false)
                                        {
                                            BubbleEvent = false;
                                            return;
                                        }
                                    }
                                    //검수입고 문서만 등록 시 주석처리_E
                                    oQEYesNo = oForm.Items.Item("QEYesNo").Specific.Value.ToString().Trim();
                                    oPurchase = oForm.Items.Item("Purchase").Specific.Value.ToString().Trim();
                                }
                                else
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            else
                            {
                                //검수입고 문서만 등록 시 주석처리_S
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                {
                                    if (PS_MM070_Add_oPurchaseDeliveryNotes() == false)
                                    {
                                        BubbleEvent = false;
                                        return;
                                    }
                                }
                                //검수입고 문서만 등록 시 주석처리_E
                            }
                        }
                        PS_MM070_Delete_EmptyRow(); //검수입고 문서만 등록 시 이 행은 주석 제외
                    }
                    else if (pVal.ItemUID == "Btn_prt")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PS_MM070_Print_Report01);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
                        {
                            PS_MM070_FormItemEnabled();
                            PS_MM070_Initialization();
                            PS_MM070_FormClear();
                            oDS_PS_MM070H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                            PS_MM070_AddMatrixRow(0, true);

                            if (!string.IsNullOrEmpty(oQEYesNo)) //검수여부 및 품의형태 세팅 로직 수정(2013.05.23 송명규 수정)
                            {
                                oForm.Items.Item("QEYesNo").Specific.Select(oQEYesNo, SAPbouiCOM.BoSearchKey.psk_ByValue);
                            }
                            else
                            {
                                oForm.Items.Item("QEYesNo").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            }
                            if (!string.IsNullOrEmpty(oPurchase))
                            {
                                oForm.Items.Item("Purchase").Specific.Select(oPurchase, SAPbouiCOM.BoSearchKey.psk_ByValue);
                            }
                            else
                            {
                                oForm.Items.Item("Purchase").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
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
                        if (pVal.ItemUID == "CardCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "CntcCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "GADocLin")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("GADocLin").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()) || string.IsNullOrEmpty(oForm.Items.Item("QEYesNo").Specific.Value.ToString().Trim()) || string.IsNullOrEmpty(oForm.Items.Item("Purchase").Specific.Value.ToString().Trim()))
                                    {
                                        dataHelpClass.MDC_GF_Message("사업장, 품질검수여부 또는 구매구분을 먼저 선택하세요.", "E");
                                        BubbleEvent = false;
                                        return;
                                    }
                                    else
                                    {
                                        PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                        BubbleEvent = false;
                                    }
                                }
                            }
                            else if (pVal.ColUID == "ItemCode")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PS_SM010 PS_SM010 = new PS_SM010();
                                    PS_SM010.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
                                    BubbleEvent = false;
                                }
                            }
                            else if (pVal.ColUID == "WhsCode")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("WhsCode").Cells.Item(pVal.Row).Specific.Value))
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
        }

        /// <summary>
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string TeamCode;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
                        oDS_PS_MM070L.Clear();
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PS_MM070_AddMatrixRow(0, false);
                        }
                        if (oForm.Items.Item("Purchase").Specific.Value.ToString().Trim() == "30" || oForm.Items.Item("Purchase").Specific.Value.ToString().Trim() == "40" || oForm.Items.Item("Purchase").Specific.Value.ToString().Trim() == "60")
                        {
                            if (oForm.Items.Item("Purchase").Specific.Value.ToString().Trim() == "40")
                            {
                                //TeamCode = dataHelpClass.User_TeamCode();
                                //if (TeamCode == "1400" || TeamCode == "2600") //1400 : 창원사업장 품보팀, 2600 : 부산사업장 품보팀
                                //{
                                //}
                                //else
                                //{
                                //    dataHelpClass.MDC_GF_Message("외주제작품의 검수입고는 품질보증팀 담당자만 가능합니다.", "E");
                                //    oDS_PS_MM070H.SetValue("U_Purchase", 0, "%");
                                //    BubbleEvent = false;
                                //}
                            }
                        }
                        else
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Mat01" && pVal.ColUID == "BadCode1")
                    {
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
            int i;
            int sSeq;
            int sCount;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "BadCode1")
                        {
                            sCount = oMat01.Columns.Item("BadCode2").Cells.Item(pVal.Row).Specific.ValidValues.Count;
                            sSeq = sCount;
                            for (i = 1; i <= sCount; i++)
                            {
                                oMat01.Columns.Item("BadCode2").Cells.Item(pVal.Row).Specific.ValidValues.Remove(sSeq - 1, SAPbouiCOM.BoSearchKey.psk_Index);
                                sSeq -= 1;
                            }

                            sQry = "SELECT Distinct Convert(int, b.U_SmalCode) As U_SmalCode, b.U_SmalName From [@PS_PP003H] a Inner Join [@PS_PP003L] b On a.DocEntry = b.DocEntry ";
                            sQry += "Where a.U_BigCode = '1' And a.U_MidCode = '" + oMat01.Columns.Item("BadCode1").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "' Order by Convert(int, b.U_SmalCode) ";
                            oRecordSet01.DoQuery(sQry);
                            while (!oRecordSet01.EoF)
                            {
                                oMat01.Columns.Item("BadCode2").ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                                oRecordSet01.MoveNext();
                            }
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "CardCode")
                        {
                            PS_MM070_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "CntcCode")
                        {
                            PS_MM070_FlushToItemValue(pVal.ItemUID, 0, "");
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "GADocLin")
                            {
                                PS_MM070_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                                if (oErrNum == 1)
                                {
                                    oErrNum = 0;
                                    BubbleEvent = false;
                                }
                            }
                            else if (pVal.ColUID == "ItemCode")
                            {
                                PS_MM070_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "WhsCode")
                            {
                                PS_MM070_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ColUID == "InQty" || pVal.ColUID == "Qty" || pVal.ColUID == "RealWt" || pVal.ColUID == "Weight")
                            {
                                PS_MM070_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
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
                    PS_MM070_AddMatrixRow(oMat01.RowCount, false);
                    PS_MM070_FormItemEnabled();
                    oMat01.AutoResizeColumns();
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
            string GADocLin;

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01" && pVal.ColUID == "GADocLin")
                    {
                        GADocLin = oMat01.Columns.Item("GADocLin").Cells.Item(pVal.Row).Specific.Value;
                        PS_MM050 pS_MM050 = new PS_MM050();
                        pS_MM050.LoadForm(GADocLin.Substring(0, GADocLin.IndexOf('-')));
                        pS_MM050 = null;
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM070H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM070L);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                        if (PS_MM070_Validate("행삭제") == false)
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
                        oDS_PS_MM070L.RemoveRecord(oDS_PS_MM070L.Size - 1);
                        oMat01.LoadFromDataSource();

                        if (oMat01.RowCount == 0)
                        {
                            PS_MM070_AddMatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_MM070L.GetValue("U_ItemName", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_MM070_AddMatrixRow(oMat01.RowCount, false);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            // 마감일자 Check
                            if (dataHelpClass.Check_Finish_Status(oDS_PS_MM070H.GetValue("U_BPLId", 0).ToString().Trim(), oDS_PS_MM070H.GetValue("U_DocDate", 0).ToString().Trim().Substring(0, 6)) == false)
                            {
                                PSH_Globals.SBO_Application.MessageBox("마감상태가 잠금입니다. 해당 일자로 취소할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.");
                                BubbleEvent = false;
                                return;
                            }
                            if (PS_MM070_Add_oPurchaseReturns() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1286": //닫기
                            // 마감일자 Check
                            if (dataHelpClass.Check_Finish_Status(oDS_PS_MM070H.GetValue("U_BPLId", 0).ToString().Trim(), oDS_PS_MM070H.GetValue("U_DocDate", 0).ToString().Trim().Substring(0, 6)) == false)
                            {
                                PSH_Globals.SBO_Application.MessageBox("마감상태가 잠금입니다. 해당 일자로 닫기할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.");
                                BubbleEvent = false;
                                return;
                            }
                            if (PS_MM070_Close_oPurchaseDeliveryNotes() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            PS_MM070_FormItemEnabled();
                            oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            PS_MM070_FormItemEnabled();
                            oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_MM070_Initialization();
                            PS_MM070_FormClear();
                            PS_MM070_AddMatrixRow(0, true);
                            oDS_PS_MM070H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
                            PS_MM070_FormItemEnabled();
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
                            PS_MM070_FormItemEnabled();
                            if (oMat01.VisualRowCount > 0)
                            {
                                if (!string.IsNullOrEmpty(oMat01.Columns.Item("GADocLin").Cells.Item(oMat01.VisualRowCount).Specific.Value))
                                {
                                    if (oDS_PS_MM070H.GetValue("Status", 0) == "O")
                                    {
                                        PS_MM070_AddMatrixRow(oMat01.RowCount, false);
                                    }
                                }
                            }
                            break;
                        case "1287": //복제
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
