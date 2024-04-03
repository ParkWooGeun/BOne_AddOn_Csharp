using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using System.Collections.Generic;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 1.분말패킹등록
    /// </summary>
    internal class PS_PP092 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_PP092H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_PP092L; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private string sDocNum;
        private string sPackNo;
        private string Last_CntcCode;
        private string Last_CntcName;
        private string Last_InDate;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP092.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP092_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP092");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_PP092_CreateItems();
                PS_PP092_ComboBox_Setting();
                PS_PP092_EnableMenus();
                PS_PP092_SetDocument(oFormDocEntry);
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
        private void PS_PP092_CreateItems()
        {
            try
            {
                oDS_PS_PP092H = oForm.DataSources.DBDataSources.Item("@PS_PP092H");
                oDS_PS_PP092L = oForm.DataSources.DBDataSources.Item("@PS_PP092L");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oForm.Items.Item("InDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

                oForm.DataSources.UserDataSources.Add("SumWeight", SAPbouiCOM.BoDataType.dt_QUANTITY, 10);
                oForm.Items.Item("SumWeight").Specific.DataBind.SetBound(true, "", "SumWeight");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP092_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT TOP 2 BPLId, BPLName FROM [OBPL]  order by BPLId", "1", false, false);
                dataHelpClass.Combo_ValidValues_Insert("PS_PP092", "VIEWYN", "", "Y", "출력");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP092", "VIEWYN", "", "N", "미출력");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("VIEWYN").Specific, "PS_PP092", "VIEWYN", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP092", "ShipType", "", "1", "화물");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP092", "ShipType", "", "2", "택배");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("ShipType").Specific, "PS_PP092", "ShipType", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP092", "CheckYN", "", "N", "미확인");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP092", "CheckYN", "", "Y", "확인");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("CheckYN").Specific, "PS_PP092", "CheckYN", false);
                oForm.Items.Item("CheckYN").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);

                oMat01.Columns.Item("OutGbn").ValidValues.Add("10", "판매출고");
                oMat01.Columns.Item("OutGbn").ValidValues.Add("20", "샘플출고");
                oMat01.Columns.Item("OutGbn").ValidValues.Add("30", "무상출고");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PS_PP092_Validate
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PS_PP092_Validate(string ValidateType)
        {
            bool returnValue = false;
            string Query01;
            string Query02;
            string Query03;
            string Query04;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet03 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet04 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
                            Query01 = " Select sum(Quantity) as Qty FROM OIBT ";
                            Query01 += " WHERE ItemCode='" + oMat01.Columns.Item("ItemCode").Cells.Item(oLastColRow01).Specific.Value + "'";
                            Query01 += "       AND BatchNum='" + oMat01.Columns.Item("LotNo").Cells.Item(oLastColRow01).Specific.Value + "'";
                            oRecordSet01.DoQuery(Query01);

                            if (oRecordSet01.Fields.Item("Qty").Value <= 0)
                            {
                                errMessage = "출하처리된 품목입니다. 삭제할수 없습니다.";
                                throw new Exception();
                            }
                            Query01 = " SELECT Count(*) as cnt ";
                            Query01 += "  FROM [@PS_SD040H] a inner join [@PS_SD040L] b on a.DocEntry = b.DocEntry and a.Canceled ='N'";
                            Query01 += " WHERE b.U_PackNo ='" + oForm.Items.Item("PackNo").Specific.Value + "'";
                            oRecordSet01.DoQuery(Query01);

                            if (oRecordSet01.Fields.Item(0).Value > 0)
                            {
                                errMessage = "납품처리된 패킹번호입니다. 취소할수 없습니다.";
                                throw new Exception();
                            }
                        }
                    }
                }
                else if (ValidateType == "취소")
                {
                    if (oForm.Items.Item("Canceled").Specific.Value == "Y")
                    {
                        errMessage = "이미취소된문서입니다.";
                        throw new Exception();
                    }
                    Query01 = " SELECT Count(*) ";
                    Query01 += "  FROM [@PS_SD040H] a inner join [@PS_SD040L] b on a.DocEntry = b.DocEntry and a.Canceled ='N'";
                    Query01 += " WHERE b.U_PackNo ='" + oForm.Items.Item("PackNo").Specific.Value + "'";
                    oRecordSet01.DoQuery(Query01);

                    if (oRecordSet01.Fields.Item(0).Value > 0)
                    {
                        errMessage = "납품처리된 패킹번호입니다. 취소할수 없습니다.";
                        throw new Exception();
                    }

                    Query01 = "SELECT U_LotNo,U_ItemCode,U_PackNo FROM [@PS_PP092L] WHERE  DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                    oRecordSet01.DoQuery(Query01);

                    while (oRecordSet01.EoF == false)
                    {
                        Query02 = "SELECT Sum(Quantity) as Qty FROM OIBT ";
                        Query02 += " WHERE ItemCode = '" + oRecordSet01.Fields.Item("U_ItemCode").Value + "'";
                        Query02 += "       AND BatchNum='" + oRecordSet01.Fields.Item("U_LotNo").Value + "'";
                        oRecordSet02.DoQuery(Query02);


                        if (oRecordSet02.Fields.Item("Qty").Value <= 0) //멀티는 일부가나갈수가 없다 수량이 없다면 출고된것으로 본다.
                        {
                            errMessage = "이미 출고된 품목이 있습니다. 취소할 수 없습니다.";
                            throw new Exception();

                        }
                        else //재고가 다 있다면 행별로 OBTN에 저장했던PackNo를 지워줘야한다.
                        {
                            Query03 = "SELECT U_LotNo,U_ItemCode,U_PackNo FROM [@PS_PP092L] WHERE  DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                            oRecordSet03.DoQuery(Query03);

                            while (oRecordSet03.EoF == false)
                            {
                                Query04 = "UPDATE OBTN SET U_PackNo=''";
                                Query04 += " WHERE ItemCode = '" + oRecordSet01.Fields.Item("U_ItemCode").Value + "'";
                                Query04 += "       AND DistNumber='" + oRecordSet01.Fields.Item("U_LotNo").Value + "'";

                                oRecordSet04.DoQuery(Query04);
                                oRecordSet03.MoveNext();
                            }
                        }

                        oRecordSet01.MoveNext();
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
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet03);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet04);
            }
            return returnValue;
        }

        /// <summary>
        /// UpdateToPP092
        /// </summary>
        /// <param name="sPackNum"></param>
        /// <param name="sDocNum"></param>
        /// <returns></returns>
        private bool PS_PP092_UpdateToPP092(string sPackNum, string sDocNum)
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            string Query01;
            string Query02;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                Query01 = "Select U_ItemCode,U_LotNo from [@PS_PP092L] WHERE DocEntry='" + sDocNum + "' and U_PackNo= '" + sPackNum.ToString().Trim() + "'";
                oRecordSet01.DoQuery(Query01);

                //아이템과 LOT번호는 고유함
                while (oRecordSet01.EoF == false)
                {
                    Query02 = "UPDATE [OBTN] SET ";
                    Query02 += " U_PackNo='" + sPackNum.ToString().Trim() + "'";
                    Query02 += "  Where ItemCode = '" + oRecordSet01.Fields.Item("U_ItemCode").Value + "'";
                    Query02 += "       And DistNumber= '" + oRecordSet01.Fields.Item("U_LotNo").Value + "'";
                    oRecordSet02.DoQuery(Query02);

                    oRecordSet01.MoveNext();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
            }
            return returnValue;
        }

        /// <summary>
        /// Calc_SumWeight
        /// </summary>
        /// <returns></returns>
        private void PS_PP092_Calc_SumWeight()
        {
            int i;
            double SumWeight = 0;

            try
            {
                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    SumWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i).Specific.Value);
                }
                oForm.Items.Item("SumWeight").Specific.Value = SumWeight;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_PP092_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, true, false, false, false, false, false, false);
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
        private void PS_PP092_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_PP092_FormItemEnabled();
                    PS_PP092_AddMatrixRow(0, true);
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
        private void PS_PP092_FormItemEnabled()
        {
            string Query01;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP092_FormClear();
                    oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index); //콤보기본선택
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                    oForm.Items.Item("InDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                    oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
                    oForm.Items.Item("empty").Click();
                    oForm.Items.Item("Mat01").Enabled = true; //메트릭스
                    oForm.Items.Item("CntcCode").Enabled = true;  //작성자
                    oForm.Items.Item("InDate").Enabled = true; //작성일
                    oForm.Items.Item("BPLId").Enabled = true; //사업장
                    oForm.Items.Item("DocEntry").Enabled = false; //문서번호
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("DocEntry").Enabled = true; //문서번호활성화
                    oForm.Items.Item("InDate").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    Query01 = "Select Distinct Quantity FROM OIBT WHERE BatchNum = '" + oMat01.Columns.Item("LotNo").Cells.Item(1).Specific.Value + "'";
                    oRecordSet01.DoQuery(Query01);

                    if (oDS_PS_PP092H.GetValue("Canceled", 0) == "Y" || oRecordSet01.Fields.Item(0).Value < 0)
                    {
                        oForm.Items.Item("Mat01").Enabled = false;
                        oForm.Items.Item("CntcCode").Enabled = false;
                        oForm.Items.Item("DocEntry").Enabled = false;
                        oForm.Items.Item("InDate").Enabled = false;
                        oForm.Items.Item("BPLId").Enabled = false;
                    }
                    else
                    {
                        oForm.Items.Item("DocEntry").Enabled = false; //찾기하고나면 문서비활성화처리
                        oForm.Items.Item("CntcCode").Enabled = true;
                        oForm.Items.Item("Mat01").Enabled = true;
                        oForm.Items.Item("InDate").Enabled = true;
                        oForm.Items.Item("BPLId").Enabled = true;
                    }
                    if (oForm.Items.Item("CheckYN").Specific.Value.ToString().Trim() == "Y")
                    {
                        oForm.Items.Item("Mat01").Enabled = false;
                        oForm.Items.Item("CntcCode").Enabled = false;
                        oForm.Items.Item("DocEntry").Enabled = false;
                        oForm.Items.Item("InDate").Enabled = false;
                        oForm.Items.Item("BPLId").Enabled = false;
                        oForm.Items.Item("SD030Num").Enabled = false;
                        oForm.Items.Item("CardCode").Enabled = false;
                        oForm.Items.Item("Destinaton").Enabled = false;
                        oForm.Items.Item("ShipType").Enabled = false;
                        oForm.Items.Item("Comments").Enabled = false;
                        oForm.Items.Item("1").Enabled = false;
                    }
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
        /// PS_PP092_AddMatrixRow
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_PP092_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_PP092L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_PP092L.Offset = oRow;
                oDS_PS_PP092L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        /// PS_PP092_MTX01
        /// </summary>
        private void PS_PP092_MTX01()
        {
            string errMessage = string.Empty;
            int i;
            string Query01;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                Param01 = oForm.Items.Item("Param01").Specific.Value.ToString().Trim();
                Param02 = oForm.Items.Item("Param01").Specific.Value.ToString().Trim();
                Param03 = oForm.Items.Item("Param01").Specific.Value.ToString().Trim();
                Param04 = oForm.Items.Item("Param01").Specific.Value.ToString().Trim();

                Query01 = "SELECT 10";
                oRecordSet01.DoQuery(Query01);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }
                ProgressBar01.Text = "조회시작";

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_PP092L.InsertRecord(i);
                    }
                    oDS_PS_PP092L.Offset = i;
                    oDS_PS_PP092L.SetValue("U_COL01", i, oRecordSet01.Fields.Item(0).Value);
                    oDS_PS_PP092L.SetValue("U_COL02", i, oRecordSet01.Fields.Item(1).Value);
                    oRecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
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
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_PP092_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP092'", "");
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
        /// 문서 취소시 실행
        /// </summary>
        /// <param name="PackNo"></param>
        /// <returns></returns>
        private bool PS_PP092_CancelPackNo(string PackNo)
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            string Query01;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                Query01 = "UPDATE Z_PACKING_LOT SET BARCDYN ='N', PACKNO ='', CheckDate ='29991231' WHERE PACKNO ='" + PackNo + "'";
                oRecordSet01.DoQuery(Query01);
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
            }
            return returnValue;
        }


        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PS_PP092_DataValidCheck()
        {
            bool returnValue = false;
            int i = 0;
            string errMessage = string.Empty;
            string ClickCode = string.Empty;
            string type = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("InDate").Specific.Value))
                {
                    errMessage = "작성일은 필수입니다.";
                    ClickCode = "InDate";
                    type = "F";
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value))
                {
                    errMessage = "거래처는 필수입니다.";
                    ClickCode = "CardCode";
                    type = "F";
                    oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
                    //값이 한줄들어있을때 한줄삭제후 갱신한다거나한다면
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("LotNo").Cells.Item(1).Specific.Value))
                    {
                        errMessage = "Matrix값이 한줄이상은 있어야합니다.";
                        type = "X";
                        throw new Exception();
                    }
                }
                for (i = 1; i <= (oMat01.VisualRowCount - 1); i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("OutGbn").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "출고구분은 필수입니다.";
                        ClickCode = "OutGbn";
                        type = "M";
                        throw new Exception();
                    }
                    else
                    {
                        if (oMat01.Columns.Item("OutGbn").Cells.Item(i).Specific.Value == "10")
                        {
                            if (string.IsNullOrEmpty(oMat01.Columns.Item("SD030Num").Cells.Item(i).Specific.Value))
                            {
                                errMessage = "출하요청번호를 입력하셔야 합니다.";
                                ClickCode = "SD030Num";
                                type = "M";
                                throw new Exception();
                            }
                        }
                    }
                }
                oDS_PS_PP092L.RemoveRecord(oDS_PS_PP092L.Size - 1);
                oMat01.LoadFromDataSource();
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP092_FormClear();
                }
                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    if (type == "F")
                    {
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                        oForm.Items.Item(ClickCode).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                    else if (type == "M")
                    {
                        PSH_Globals.SBO_Application.MessageBox(errMessage);
                        oMat01.Columns.Item(ClickCode).Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
        /// 리포트 조회
        /// </summary>
        [STAThread]
        private void PS_PP092_Print_Report01()
        {
            string WinTitle;
            string ReportName;
            string DocEntry;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

                WinTitle = "[PS_QM008_10] 검사성적서 출력(한글)";
                ReportName = "PS_QM008_10.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>
                {
                    new PSH_DataPackClass("@DocEntry", DocEntry),
                    new PSH_DataPackClass("@VIEWYN", "Y"),
                    new PSH_DataPackClass("@SOVIEWYN", "Y"),
                    new PSH_DataPackClass("@Gubun", "P"),
                    new PSH_DataPackClass("@Lang", "E")
                }; //Parameter

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 리포트 조회
        /// </summary>
        [STAThread]
        private void PS_PP092_Print_Report02()
        {
            string WinTitle;
            string ReportName;
            string DocEntry;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

                WinTitle = "[PS_QM008_30] 검사성적서 출력(영문)";
                ReportName = "PS_QM008_30.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>
                {
                    new PSH_DataPackClass("@DocEntry", DocEntry),
                    new PSH_DataPackClass("@VIEWYN", "Y"),
                    new PSH_DataPackClass("@SOVIEWYN", "Y"),
                    new PSH_DataPackClass("@Gubun", "P"),
                    new PSH_DataPackClass("@Lang", "E")
                }; //Parameter

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 리포트 조회
        /// </summary>
        [STAThread]
        private void PS_PP092_Print_Report03()
        {
            string WinTitle;
            string ReportName;
            string DocEntry;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

                WinTitle = "[PS_QM008_20] 검사성적서 출력(중문)";
                ReportName = "PS_QM008_20.rpt";
                //Parameter
                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>
                {
                    new PSH_DataPackClass("@DocEntry", DocEntry),
                    new PSH_DataPackClass("@VIEWYN", "Y"),
                    new PSH_DataPackClass("@SOVIEWYN", "Y"),
                    new PSH_DataPackClass("@Gubun", "P"),
                    new PSH_DataPackClass("@Lang", "C")
                };

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        ///// <summary>
        ///// 리포트 조회
        ///// </summary>
        //[STAThread]
        //private void PS_PP092_Print_Report04()
        //{
        //    string WinTitle;
        //    string ReportName;
        //    string Param01;
        //    string Param02;
        //    string Param03;
        //    string Param04;
        //    string Param05;
        //    string Param06;
        //    string Param07;
        //    PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

        //    try
        //    {
        //        Param01 = oForm.Items.Item("BPLId").Specific.Selected.Value;
        //        Param02 = oForm.Items.Item("DocDate").Specific.Value;
        //        Param03 = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
        //        Param04 = oForm.Items.Item("ItemName").Specific.Value;
        //        Param05 = oForm.Items.Item("OrdNum").Specific.Value;
        //        Param06 = oForm.Items.Item("BatchNum").Specific.Value;
        //        Param07 = oForm.Items.Item("CardCode").Specific.Value;

        //        WinTitle = "BOX-LABEL출력[PS_PACKING_PD_05] ";
        //        ReportName = "PS_PACKING_PD_05.rpt";
        //        //Parameter
        //        List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>
        //        {
        //            new PSH_DataPackClass("@BPLId", Param01),
        //            new PSH_DataPackClass("@DocDate", Param02),
        //            new PSH_DataPackClass("@ItemCode", Param03),
        //            new PSH_DataPackClass("@ItemName", Param04),
        //            new PSH_DataPackClass("@OrdNum", Param05),
        //            new PSH_DataPackClass("@BatchNum", Param06),
        //            new PSH_DataPackClass("@CardCode", Param07)
        //        };

        //        formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, "Y");
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //}

        /// <summary>
        /// 리포트 조회
        /// </summary>
        [STAThread]
        private void PS_PP092_Print_Report05()
        {
            string WinTitle;
            string ReportName;
            string Param01;
            string Param02;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                Param01 = oForm.Items.Item("BPLId").Specific.Selected.Value;
                Param02 = oForm.Items.Item("DocEntry").Specific.Value;

                WinTitle = "PACKING LIST출력[PS_PP092_13] ";
                ReportName = "PS_PP092_13.rpt";

                //Parameter
                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>
                {
                    new PSH_DataPackClass("@BPLId", Param01),
                    new PSH_DataPackClass("@DocEntry", Param02)
                };

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, "Y");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            int i;
            string Query01;
            string sPackDate;
            string sIndex;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_PP092_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            //해야할일 작업 (Pack번호를 순번해서 이문서에 업데이트해준다.)
                            sPackDate = DateTime.Now.ToString("yyyyMMdd");

                            Query01 = "Select ISNULL(MAX(CONVERT(INT,RIGHT(U_PackNo,2))),0)+1 FROM [@PS_PP092H] WHERE Canceled = 'N' and CreateDate = '" + sPackDate + "'";
                            oRecordSet01.DoQuery(Query01);

                            sIndex = oRecordSet01.Fields.Item(0).Value.ToString();
                            sPackNo = sPackDate + sIndex.PadLeft(2, '0');

                            oDS_PS_PP092H.SetValue("U_PackNo", 0, sPackNo);

                            for (i = 0; i <= (oMat01.VisualRowCount - 1); i++)
                            {
                                oDS_PS_PP092L.SetValue("U_PackNo", i, sPackNo);
                                oMat01.Columns.Item("PackNo").Cells.Item(i + 1).Specific.Value = sPackNo;
                            }

                            for (i = 1; i <= oMat01.VisualRowCount; i++)
                            {
                                string BatchNumCheck;
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("LotNoSub").Cells.Item(i).Specific.Value.ToString().Trim()))
                                {
                                    BatchNumCheck = oMat01.Columns.Item("LotNo").Cells.Item(i).Specific.Value;
                                }
                                else
                                {
                                    BatchNumCheck = oMat01.Columns.Item("LotNoSub").Cells.Item(i).Specific.Value;
                                }
                                Query01 = "UPDATE Z_PACKING_LOT SET PackNo ='" + sPackNo + "' where BarCDYN ='N' and BatchNum = '" + BatchNumCheck + "'";
                                oRecordSet01.DoQuery(Query01);
                            }

                            //BeforeAction 이 false가될때 OBTN에도 PACKNO정보를 행별품목에 업뎃해주어야함
                            sDocNum = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                            //문서번호 전역변수에 담음

                            //추가완료 후 다시 자동으로 CnctCode,CnctName,InDate를 보여주기 위해 미리 저장
                            Last_CntcCode = oDS_PS_PP092H.GetValue("U_CntcCode", 0).ToString().Trim();
                            Last_CntcName = oDS_PS_PP092H.GetValue("U_CntcName", 0).ToString().Trim();
                            Last_InDate = oDS_PS_PP092H.GetValue("U_InDate", 0).ToString().Trim();

                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_PP092_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            for (i = 1; i <= oMat01.VisualRowCount; i++)
                            {
                                string BatchNumCheck;
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("LotNoSub").Cells.Item(i).Specific.Value.ToString().Trim()))
                                {
                                    BatchNumCheck = oMat01.Columns.Item("LotNo").Cells.Item(i).Specific.Value;
                                }
                                else
                                {
                                    BatchNumCheck = oMat01.Columns.Item("LotNoSub").Cells.Item(i).Specific.Value;
                                }
                                Query01 = "UPDATE Z_PACKING_LOT SET PackNo ='" + oForm.Items.Item("PackNo").Specific.Value + "' where BarCDYN ='N' and BatchNum = '" + BatchNumCheck + "'";
                                oRecordSet01.DoQuery(Query01);
                            }
                        }
                    }
                    else if (pVal.ItemUID == "Btn_Prt")//국문
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PS_PP092_Print_Report01);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                    else if (pVal.ItemUID == "Btn_Prt2")//영문
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PS_PP092_Print_Report02);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                    else if (pVal.ItemUID == "Btn_Prt1") //한문
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PS_PP092_Print_Report03);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                    //else if (pVal.ItemUID == "Btn_Label") //라벨
                    //{
                    //    System.Threading.Thread thread = new System.Threading.Thread(PS_PP092_Print_Report04);
                    //    thread.SetApartmentState(System.Threading.ApartmentState.STA);
                    //    thread.Start();
                    //if (oLastColRow01 != 0)
                    //{
                    //    PS_PP092S pS_PP092S = new PS_PP092S();
                    //    string PackNo;
                    //    string InspNo;
                    //    string ProDate;
                    //    PackNo = oForm.Items.Item("PackNo").Specific.Value.ToString().Trim();
                    //    InspNo = oMat01.Columns.Item("InspNo").Cells.Item(oLastColRow01).Specific.Value.ToString().Trim();
                    //    ProDate = oMat01.Columns.Item("ProDate").Cells.Item(oLastColRow01).Specific.Value;
                    //    pS_PP092S.LoadForm(PackNo, InspNo, ProDate);
                    //    BubbleEvent = false;
                    //}
                    //else
                    //{
                    //    errMessage = "출력할 행을 선택 후 LABEL 출력을 누르세요.";
                    //    throw new Exception();
                    //}
                    //}
                    else if (pVal.ItemUID == "Btn_Pack") //라벨_Packing list 출력
                    {

                        System.Threading.Thread thread = new System.Threading.Thread(PS_PP092_Print_Report05);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
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
                                PS_PP092_FormItemEnabled();
                                PS_PP092_UpdateToPP092(sPackNo, sDocNum);
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                PSH_Globals.SBO_Application.ActivateMenuItem("1291");
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_PP092_FormItemEnabled();
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
                        if (pVal.ItemUID == "CntcCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        if (pVal.ItemUID == "SD030Num")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("SD030Num").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "UniqueK"); //Lot번호
                        dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "LotNo"); //Lot번호
                        dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "SD030Num"); //출하요청번호
                        dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "InspNo"); //검사번호
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
            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    oMat01.FlushToDataSource();
                    oMat01.LoadFromDataSource();
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
                            oLastColRow01 = pVal.Row;
                        }
                    }
                    else if (pVal.ItemUID == "1")
                    {
                        if (pVal.ItemUID == "1")
                        {
                            oForm.EnableMenu("1281", true); //찾기하고 다시 찾기아이콘활성화처리
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
        }

        /// <summary>
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string SD030Num;
            string SD030H;
            string SD030L;
            string Query01;
            string Query02;
            string Query03;
            double Quantity;
            string BatchNum;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet03 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "UniqueK")
                            {
                                BatchNum = oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.Split('_')[0];
                                oDS_PS_PP092L.SetValue("U_LotNo", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.Split('_')[0]);
                                oDS_PS_PP092L.SetValue("U_LotNoSub", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.Split('_')[1]);

                                //기타작업
                                SD030Num = oForm.Items.Item("SD030Num").Specific.Value;
                                SD030H = oForm.Items.Item("SD030H").Specific.Value;
                                SD030L = oForm.Items.Item("SD030L").Specific.Value;
                                if (!string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.Split('_')[1]))
                                {

                                    BatchNum = oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.Split('_')[1];

                                    Query01 = " select a.ItemCode as ItemCode";
                                    Query01 += "     , a.ItemName as ItemName";
                                    Query01 += " 	 , b.Quantity as Quantity";
                                    Query01 += " 	 , b.Quantity as GQty";
                                    Query01 += " 	 , c.u_weight as PQty";
                                    Query01 += " 	 , a.InDate as CreateDate";
                                    Query01 += " 	 , case when isnull(b.bsinspno,'') = '' then b.InspNo else b.bsinspno end as InspNo";
                                    Query01 += " 	 , b.CardSeq as CardSeq";
                                    Query01 += " From OIBT a Inner Join Z_PACKING_PD b On a.ItemCode = b.ItemCode And a.BatchNum = (case when isnull(b.bBatchNum,'') = '' then b.BatchNum else b.bBatchNum end) ";
                                    Query01 += " 			 left join (select u_lotno, sum(u_weight)as u_weight from [@PS_PP092l] where 1=1 group by u_lotno) c On a.BatchNum = c.u_lotno";
                                    Query01 += "  Where a.Quantity > 0 ";
                                    Query01 += "  And a.BatchNum = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.Split('_')[0] + "'";
                                    Query01 += " And b.BatchNum = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.Split('_')[1] + "'";
                                }
                                else
                                {
                                    Query01 = " select a.ItemCode as ItemCode";
                                    Query01 += "      , a.ItemName as ItemName";
                                    Query01 += " 	 , a.Quantity as Quantity";
                                    Query01 += " 	 , a.Quantity as GQty";
                                    Query01 += " 	 , c.u_weight as PQty";
                                    Query01 += " 	 , a.InDate as CreateDate";
                                    Query01 += " 	 , case when isnull(b.bsinspno,'') = '' then b.InspNo else b.bsinspno end as InspNo";
                                    Query01 += " 	 , b.CardSeq as CardSeq";
                                    Query01 += " From OIBT a Inner Join Z_PACKING_PD b On a.ItemCode = b.ItemCode And a.BatchNum = (case when isnull(b.bBatchNum,'') = '' then b.BatchNum else b.bBatchNum end) ";
                                    Query01 += " 			 left join (select u_lotno, sum(u_weight)as u_weight from [@PS_PP092l] where 1=1 group by u_lotno) c On a.BatchNum = c.u_lotno";
                                    Query01 += "  Where a.Quantity > 0 ";
                                    Query01 += "  And a.BatchNum = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.Split('_')[0] + "'";
                                }
                                oRecordSet01.DoQuery(Query01);

                                Query02 = "Select UnitWgt = a.U_RelCd From [@PS_SY001L] a Where a.Code = 'P204' and Left( '" + BatchNum + "' ,1) = a.U_Minor";
                                oRecordSet02.DoQuery(Query02);

                                Query03 = "Select Cnt = Count(a.BatchNum) From Z_PACKING_PD a Where Case When Isnull(a.bBatchNum,'') = '' Then a.BatchNum Else a.bBatchNum End = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.Split('_')[1] + "'";
                                oRecordSet03.DoQuery(Query03);

                                if (oRecordSet01.RecordCount == 0)
                                {
                                    oDS_PS_PP092L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                }
                                else
                                {
                                    oDS_PS_PP092L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);

                                    if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP092L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                    {
                                        PS_PP092_AddMatrixRow(pVal.Row, false);
                                    }
                                    oDS_PS_PP092L.SetValue("U_SD030Num", pVal.Row - 1, SD030Num);
                                    oDS_PS_PP092L.SetValue("U_ItemCode", pVal.Row - 1, oRecordSet01.Fields.Item("ItemCode").Value);
                                    oDS_PS_PP092L.SetValue("U_ItemName", pVal.Row - 1, oRecordSet01.Fields.Item("ItemName").Value);
                                    oDS_PS_PP092L.SetValue("U_Quantity", pVal.Row - 1, oRecordSet01.Fields.Item("Quantity").Value);
                                    oDS_PS_PP092L.SetValue("U_PQty", pVal.Row - 1, oRecordSet01.Fields.Item("PQty").Value);
                                    oDS_PS_PP092L.SetValue("U_GQty", pVal.Row - 1, oRecordSet01.Fields.Item("GQty").Value);
                                    oDS_PS_PP092L.SetValue("U_Qty", pVal.Row - 1, Convert.ToString(System.Math.Round(oRecordSet01.Fields.Item("Quantity").Value / Convert.ToInt32(oRecordSet02.Fields.Item("UnitWgt").Value), 0)));
                                    oDS_PS_PP092L.SetValue("U_UnitWgt", pVal.Row - 1, Convert.ToString(Convert.ToInt32(oRecordSet02.Fields.Item("UnitWgt").Value)));
                                    oDS_PS_PP092L.SetValue("U_Weight", pVal.Row - 1, oRecordSet01.Fields.Item("GQty").Value);
                                    oDS_PS_PP092L.SetValue("U_ProDate", pVal.Row - 1, oRecordSet01.Fields.Item("CreateDate").Value.ToString("yyyyMMdd"));
                                    oDS_PS_PP092L.SetValue("U_ItmBsort", pVal.Row - 1, "111");
                                    oDS_PS_PP092L.SetValue("U_SD030H", pVal.Row - 1, SD030H);
                                    oDS_PS_PP092L.SetValue("U_SD030L", pVal.Row - 1, SD030L);

                                    if (oRecordSet03.Fields.Item("Cnt").Value == 1)
                                    {
                                        oDS_PS_PP092L.SetValue("U_InspNo", pVal.Row - 1, oRecordSet01.Fields.Item("InspNo").Value);
                                        oDS_PS_PP092L.SetValue("U_CardSeq", pVal.Row - 1, oRecordSet01.Fields.Item("CardSeq").Value);
                                    }
                                }
                            }
                            else if (pVal.ColUID == "SD030Num")
                            {
                                oDS_PS_PP092L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);

                                Query01 = "Select DocEntry, LineId From [@PS_SD030L] ";
                                Query01 += " Where CONVERT(VARCHAR,DocEntry) + '-' + CONVERT(VARCHAR,LineId) = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'";
                                oRecordSet01.DoQuery(Query01);

                                if (oRecordSet01.RecordCount == 0)
                                {
                                    oDS_PS_PP092L.SetValue("U_SD030H", pVal.Row - 1, "");
                                    oDS_PS_PP092L.SetValue("U_SD030L", pVal.Row - 1, "");
                                }
                                else
                                {
                                    oDS_PS_PP092L.SetValue("U_SD030H", pVal.Row - 1, oRecordSet01.Fields.Item("DocEntry").Value);
                                    oDS_PS_PP092L.SetValue("U_SD030L", pVal.Row - 1, oRecordSet01.Fields.Item("LineId").Value);
                                }
                            }
                            else if (pVal.ColUID == "Qty")
                            {
                                if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) > 0)
                                {
                                    Quantity = Convert.ToDouble(oMat01.Columns.Item("Quantity").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP092L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    if (Convert.ToDouble(oMat01.Columns.Item("UnitWgt").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value) <= Quantity)
                                    {
                                        oDS_PS_PP092L.SetValue("U_Weight", pVal.Row - 1, Convert.ToDouble(oMat01.Columns.Item("UnitWgt").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value));
                                    }
                                    else
                                    {
                                        //현재수량보다 초과할 수 없음
                                        PSH_Globals.SBO_Application.MessageBox("재고보다 많이 입력할 수 없습니다.");
                                        oDS_PS_PP092L.SetValue("U_Qty", pVal.Row - 1, Convert.ToString(System.Math.Round(Quantity / oMat01.Columns.Item("UnitWgt").Cells.Item(pVal.Row).Specific.Value, 0)));
                                        oDS_PS_PP092L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(Quantity));
                                    }
                                }
                            }
                            else if (pVal.ColUID == "InspNo")
                            {
                                Query01 = "Select CardSeq = U_CardSeq From [@PS_QM008H] ";
                                Query01 += " Where Canceled = 'N' And U_InspNo = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'";
                                oRecordSet01.DoQuery(Query01);
                                if (oRecordSet01.RecordCount > 0)
                                {
                                    oDS_PS_PP092L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP092L.SetValue("U_CardSeq", pVal.Row - 1, oRecordSet01.Fields.Item("CardSeq").Value);
                                }
                            }
                            else
                            {
                                oDS_PS_PP092L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_PP092H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "CardCode")
                            {
                                oDS_PS_PP092H.SetValue("U_CardName", 0, dataHelpClass.Get_ReData("CardName", "CardCode", "[OCRD]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", ""));
                            }
                            else if (pVal.ItemUID == "CntcCode")
                            {
                                oDS_PS_PP092H.SetValue("U_CntcName", 0, dataHelpClass.Get_ReData("U_FULLNAME", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'", ""));
                            }
                            else if (pVal.ItemUID == "SD030Num")
                            {
                                Query01 = "Select ItemCode = a.U_ItemCode, ItemName = a.U_ItemName, a.DocEntry, a.LineId , Weight = a.u_Weight, Gdate = Convert(Char(8),b.u_Docdate,112) From [@PS_SD030L] a inner join [@PS_SD030H] b on a.docentry = b.docentry  ";
                                Query01 += " Where CONVERT(VARCHAR,a.DocEntry) + '-' + CONVERT(VARCHAR,a.LineId) = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'";
                                oRecordSet01.DoQuery(Query01);

                                if (oRecordSet01.RecordCount == 0)
                                {
                                    oForm.Items.Item("ItemCode").Specific.Value = "";
                                    oForm.Items.Item("ItemName").Specific.Value = "";
                                    oForm.Items.Item("SD030H").Specific.Value = "";
                                    oForm.Items.Item("SD030L").Specific.Value = "";
                                    oForm.Items.Item("Weight").Specific.Value = "";
                                    oForm.Items.Item("Gdate").Specific.Value = "";
                                    oForm.Items.Item("SD030Num").Specific.Value = "";
                                }
                                else
                                {
                                    oDS_PS_PP092H.SetValue("U_ItemCode", 0, oRecordSet01.Fields.Item("ItemCode").Value);
                                    oDS_PS_PP092H.SetValue("U_ItemName", 0, oRecordSet01.Fields.Item("ItemName").Value);
                                    oDS_PS_PP092H.SetValue("U_SD030H", 0, oRecordSet01.Fields.Item("DocEntry").Value);
                                    oDS_PS_PP092H.SetValue("U_SD030L", 0, oRecordSet01.Fields.Item("LineId").Value);
                                    oDS_PS_PP092H.SetValue("U_Weight", 0, oRecordSet01.Fields.Item("Weight").Value);
                                    oDS_PS_PP092H.SetValue("U_Gdate", 0, oRecordSet01.Fields.Item("Gdate").Value);
                                }
                            }
                            else
                            {
                            }
                        }
                        oMat01.LoadFromDataSource();
                        oMat01.AutoResizeColumns();
                        oForm.Update();
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if ((pVal.ColUID == "UniqueK") || (pVal.ColUID == "Qty"))
                            {
                                PS_PP092_Calc_SumWeight();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet03);
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
            double SumWeight = 0;

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    for (i = 1; i <= oMat01.VisualRowCount; i++)
                    {
                        SumWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i).Specific.Value);
                    }
                    oForm.Items.Item("SumWeight").Specific.Value = SumWeight;

                    PS_PP092_FormItemEnabled();
                    PS_PP092_AddMatrixRow(oMat01.VisualRowCount, false);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP092H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP092L);
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
            int i;

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (PS_PP092_Validate("행삭제") == false)
                    {
                        BubbleEvent = false;
                        return;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    for (i = 1; i <= oMat01.VisualRowCount; i++)
                    {
                        oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i; //행을 다시 순서대로정렬해서 행번에넣고(VisualCount값은 줄어든상태)
                    }

                    oMat01.FlushToDataSource();

                    oDS_PS_PP092L.RemoveRecord(oDS_PS_PP092L.Size - 1);

                    oMat01.LoadFromDataSource();
                    //그후 다시 데이터소스를 읽어와 화면완성을 한다.

                    //행이 없으면 한줄추가
                    if (oMat01.RowCount == 0)
                    {
                        PS_PP092_AddMatrixRow(0, false);
                    }
                    else
                    {
                        //현재행삭제한 행의PorNum값이 있는행지우면 넘어가고 없는 마지막행값지우면 한행추가
                        if (!string.IsNullOrEmpty(oDS_PS_PP092L.GetValue("U_LotNo", oMat01.RowCount - 1).ToString().Trim()))
                        {
                            PS_PP092_AddMatrixRow(oMat01.RowCount, false);
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
                                if (PS_PP092_Validate("취소") == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                                else
                                {
                                    PS_PP092_CancelPackNo(oForm.Items.Item("PackNo").Specific.Value);
                                }
                            }
                            else
                            {
                                PSH_Globals.SBO_Application.MessageBox("현재 모드에서는 취소할수 없습니다.");
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
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            PS_PP092_Calc_SumWeight();
                            break;
                        case "1281": //찾기
                        case "1282": //추가
                            PS_PP092_FormItemEnabled();
                            PS_PP092_AddMatrixRow(0, true);
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
                            PS_PP092_FormItemEnabled();
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
        }
    }
}