using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 멀티패킹등록
    /// </summary>
    internal class PS_PP090 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_PP090H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_PP090L; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        private string sPackNo;
        private string sDocNum;
        private string Last_CntcCode;
        private string Last_CntcName;
        private string Last_InDate;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP090.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP090_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP090");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_PP090_CreateItems();
                PS_PP090_ComboBox_Setting();
                PS_PP090_EnableMenus();
                PS_PP090_SetDocument(oFromDocEntry01);
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
        private void PS_PP090_CreateItems()
        {
            try
            {
                oDS_PS_PP090H = oForm.DataSources.DBDataSources.Item("@PS_PP090H");
                oDS_PS_PP090L = oForm.DataSources.DBDataSources.Item("@PS_PP090L");
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
        private void PS_PP090_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT TOP 2 BPLId, BPLName FROM [OBPL]  order by BPLId", "1", false, false);
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
        private bool PS_PP090_UpdateToPP090(string sPackNum, string sDocNum)
        {
            bool functionReturnValue = false;
            string errMessage = string.Empty; 
            string Query01;
            string Query02;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                Query01 = "";
                Query01 += " Select U_ItemCode,U_LotNo from [@PS_PP090L] WHERE DocEntry='" + sDocNum + "' and U_PackNo= '" + sPackNum.ToString().Trim() + "'";
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
            return functionReturnValue;
        }

        /// <summary>
        /// PS_PP090_Calc_SumWeight
        /// </summary>
        /// <returns></returns>
        private bool PS_PP090_Calc_SumWeight()
        {
            bool functionReturnValue = false;
            int i;
            double SumWeight = 0;
            string errMessage = string.Empty;

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
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            return functionReturnValue;
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_PP090_EnableMenus()
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
        /// <param name="oFromDocEntry01">DocEntry</param>
        private void PS_PP090_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PS_PP090_FormItemEnabled();
                    PS_PP090_AddMatrixRow(0, true);
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
        private void PS_PP090_FormItemEnabled()
        {
            string Query01;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    //각모드에따른 아이템설정
                    PS_PP090_FormClear();
                    oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index); //콤보기본선택
                    oForm.EnableMenu("1281", true);  //찾기
                    oForm.EnableMenu("1282", false); //추가
                    oForm.Items.Item("InDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                    oForm.Items.Item("empty").Click();
                    oForm.Items.Item("Mat01").Enabled = true; //활성     메트릭스
                    oForm.Items.Item("CntcCode").Enabled = true; //활성     작성자
                    oForm.Items.Item("InDate").Enabled = true; //활성     작성일
                    oForm.Items.Item("BPLId").Enabled = true; //활성     사업장
                    oForm.Items.Item("DocEntry").Enabled = false; //비활성   문서번호
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    //각모드에따른 아이템설정
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("DocEntry").Enabled = true; //문서번호활성화
                    oForm.Items.Item("InDate").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    //각모드에따른 아이템설정
                    Query01 = "";
                    Query01 = "Select Distinct Quantity FROM OIBT WHERE BatchNum = '" + oMat01.Columns.Item("LotNo").Cells.Item(1).Specific.Value + "'";
                    oRecordSet01.DoQuery(Query01);

                    if (oDS_PS_PP090H.GetValue("Canceled", 0) == "Y" || oRecordSet01.Fields.Item(0).Value < 0)
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
        private void PS_PP090_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)//행추가여부
                {
                    oDS_PS_PP090L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_PP090L.Offset = oRow;
                oDS_PS_PP090L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1)); // '필드명  '행  '기본값

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
        /// PS_PP090_MTX01
        /// </summary>
        private void PS_PP090_MTX01()
        {
            string errMessage = string.Empty;
            int i;
            string Query01;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
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
                ProgressBar01.Text = "조회시작!";

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_PP090L.InsertRecord(i);
                    }
                    oDS_PS_PP090L.Offset = i;
                    oDS_PS_PP090L.SetValue("U_COL01", i, oRecordSet01.Fields.Item(0).Value);
                    oDS_PS_PP090L.SetValue("U_COL02", i, oRecordSet01.Fields.Item(1).Value);
                    oRecordSet01.MoveNext();
                    ProgressBar01.Value = ProgressBar01.Value + 1;
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_PP090_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP090'", "");
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
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PS_PP090_DataValidCheck()
        {
            bool functionReturnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("InDate").Specific.Value))
                {
                    errMessage = "작성일은 필수입니다.";
                    throw new Exception();
                }
                if (oMat01.VisualRowCount == 0)
                {
                    errMessage = "라인이 존재하지 않습니다.";
                    throw new Exception();
                }
                else
                {
                    //값이 한줄들어있을때 한줄삭제후 갱신한다거나한다면
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("LotNo").Cells.Item(1).Specific.Value))
                    {
                        errMessage = "Matrix값이 한줄이상은 있어야합니다.";
                        throw new Exception();
                    }
                }
                oDS_PS_PP090L.RemoveRecord(oDS_PS_PP090L.Size - 1);
                oMat01.LoadFromDataSource();
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP090_FormClear();
                }
                functionReturnValue = true;
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
            }
            return functionReturnValue;
        }

        /// <summary>
        /// 처리가능한 Action인지 검사
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PS_PP090_Validate(string ValidateType)
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            string Query01; 
            string Query02;
            string Query03;
            string Query04;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
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
                    //행삭제전 행삭제가능여부검사
                    //추가,수정모드일때행삭제가능검사
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        //새로추가된 행인경우, 삭제하여도 무방하다
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

                            //현재행의 수량체크
                            Query01 = "";
                            Query01 += " Select sum(Quantity) as Qty FROM OIBT ";
                            Query01 += " WHERE ItemCode='" + oMat01.Columns.Item("ItemCode").Cells.Item(oLastColRow01).Specific.Value + "'";
                            Query01 += "       AND BatchNum='" + oMat01.Columns.Item("LotNo").Cells.Item(oLastColRow01).Specific.Value + "'";
                            oRecordSet01.DoQuery(Query01);
                            if (oRecordSet01.Fields.Item("Qty").Value <= 0)
                            {
                                errMessage = "출하처리된 품목입니다. 삭제할수 없습니다.";
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

                    Query01 = "";
                    Query01 += " SELECT U_LotNo,U_ItemCode,U_PackNo FROM [@PS_PP090L] WHERE  DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                    oRecordSet01.DoQuery(Query01);

                    while (oRecordSet01.EoF == false)
                    {
                        //해당배치의 수량이 없으면 나갔다고보고 취소가 안된다.
                        Query02 = "";
                        Query02 += " SELECT Sum(Quantity) as Qty FROM OIBT ";
                        Query02 += " WHERE ItemCode = '" + oRecordSet01.Fields.Item("U_ItemCode").Value + "'";
                        Query02 += "       AND BatchNum='" + oRecordSet01.Fields.Item("U_LotNo").Value + "'";
                        oRecordSet02.DoQuery(Query02);
                        //멀티는 일부가나갈수가 없다 수량이 없다면 출고된것으로 본다.
                        if (oRecordSet02.Fields.Item("Qty").Value <= 0)
                        {
                            errMessage = "이미출고된품목이 있슴니다. 취소할수 없습니다.";
                            throw new Exception();
                            //재고가 다 있다면 행별로 OBTN에 저장했던PackNo를 지워줘야한다.
                        }
                        else
                        {
                            Query03 = "";
                            Query03 += " SELECT U_LotNo,U_ItemCode,U_PackNo FROM [@PS_PP090L] WHERE  DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                            oRecordSet03.DoQuery(Query03);
                            while (oRecordSet03.EoF == false)
                            {
                                Query04 = "";
                                Query04 += " UPDATE OBTN SET U_PackNo=''";
                                Query04 += " WHERE ItemCode = '" + oRecordSet01.Fields.Item("U_ItemCode").Value + "'";
                                Query04 += "   AND DistNumber='" + oRecordSet01.Fields.Item("U_LotNo").Value + "'"; //배치번호
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

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
            int i;
            string sPackDate;
            string sIndex;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "PS_PP090")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_PP090_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            //해야할일 작업 (Pack번호를 순번해서 이문서에 업데이트해준다.)
                            sPackDate = oForm.Items.Item("InDate").Specific.Value.ToString().Trim();
                            sIndex = dataHelpClass.GetValue("EXEC PS_PP090_01 '" + oForm.Items.Item("InDate").Specific.Value + "'",0,1);
                            sPackNo = sPackDate + sIndex;

                            //Call oDS_PS_PP090H.setValue("U_PackNo", 0, sPackNo & MDC_PS_Common.GetValue("EXEC PS_PP090_01 '" & oForm.Items("InDate").Specific.Value & "'"))
                            oDS_PS_PP090H.SetValue("U_PackNo", 0, sPackNo);

                            for (i = 0; i <= (oMat01.VisualRowCount - 1); i++)
                            {
                                oDS_PS_PP090L.SetValue("U_PackNo", i, sPackNo);
                                oMat01.Columns.Item("PackNo").Cells.Item(i + 1).Specific.Value = sPackNo;
                            }
                            //BeforeAction 이 false가될때 OBTN에도 PACKNO정보를 행별품목에 업뎃해주어야함
                            sDocNum = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                            //문서번호 전역변수에 담음

                            //추가완료 후 다시 자동으로 CnctCode,CnctName,InDate를 보여주기 위해 미리 저장
                            Last_CntcCode = oDS_PS_PP090H.GetValue("U_CntcCode", 0).ToString().Trim();
                            Last_CntcName = oDS_PS_PP090H.GetValue("U_CntcName", 0).ToString().Trim();
                            Last_InDate = oDS_PS_PP090H.GetValue("U_InDate", 0).ToString().Trim();

                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_PP090_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            //해야할일 작업
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {

                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                        {

                        }
                        //전기버튼클릭시
                    }
                    else if (pVal.ItemUID == "Btn1")
                    {
                        if (PS_PP090_DataValidCheck() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }

                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "PS_PP090")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_PP090_FormItemEnabled();
                                PS_PP090_AddMatrixRow(0, true);
                                //UDO방식일때

                                PS_PP090_UpdateToPP090(sPackNo, sDocNum);
                                //해당LOT품에대해서 OBTN테이블에 PACK번호를 업뎃해준다.

                                oDS_PS_PP090H.SetValue("U_CntcCode", 0, Last_CntcCode);
                                oDS_PS_PP090H.SetValue("U_CntcName", 0, Last_CntcName);
                                oDS_PS_PP090H.SetValue("U_InDate", 0, Last_InDate);
                                oForm.Items.Item("SumWeight").Specific.Value = "";

                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_PP090_FormItemEnabled();
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
                    //헤더에 질의연결이벤트
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", ""); //작성자코드
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "LotNo"); //Lot번호
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
                            //메트릭스 한줄선택시 반전시켜주는 구문
                        }
                    }
                    else if (pVal.ItemUID == "1")
                    {

                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        oForm.EnableMenu("1281", true);
                        //찾기하고 다시 찾기아이콘활성화처리
                    }
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
            string Query01;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "LotNo")
                            {
                                //기타작업

                                Query01 = "Select ItemCode = Max(ItemCode), ItemName = Max(ItemName), Quantity = Sum(Quantity), CreateDate = Max(CreateDate) From OIBT ";
                                Query01 = Query01 + " Where Quantity > 0 And BatchNum = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'";
                                oRecordSet01.DoQuery(Query01);

                                if (oRecordSet01.RecordCount == 0)
                                {
                                    oDS_PS_PP090L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                }
                                else
                                {
                                    oDS_PS_PP090L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    //자기자신먼저 Flush처리

                                    if (oMat01.RowCount == pVal.Row & !string.IsNullOrEmpty(oDS_PS_PP090L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                    {
                                        PS_PP090_AddMatrixRow(pVal.Row, false);
                                    }

                                    oDS_PS_PP090L.SetValue("U_ItemCode", pVal.Row - 1, oRecordSet01.Fields.Item("ItemCode").Value);
                                    oDS_PS_PP090L.SetValue("U_ItemName", pVal.Row - 1, oRecordSet01.Fields.Item("ItemName").Value);
                                    oDS_PS_PP090L.SetValue("U_Qty", pVal.Row - 1, Convert.ToString(1));
                                    oDS_PS_PP090L.SetValue("U_Weight", pVal.Row - 1, oRecordSet01.Fields.Item("Quantity").Value);
                                    oDS_PS_PP090L.SetValue("U_ProDate", pVal.Row - 1, oRecordSet01.Fields.Item("CreateDate").Value.ToString("yyyyMMdd"));
                                    oDS_PS_PP090L.SetValue("U_ItmBsort", pVal.Row - 1, "104");
                                }
                            }
                            else
                            {
                                oDS_PS_PP090L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_PP090H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "CntcCode")
                            {
                                oDS_PS_PP090H.SetValue("U_CntcName", 0, dataHelpClass.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", ""));
                            }
                            else
                            {
                                //Call oDS_PS_PP090H.setValue("U_" & pVal.ItemUID, 0, oForm.Items(pVal.ItemUID).Specific.Value)
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
                            if (pVal.ColUID == "LotNo")
                            {
                                PS_PP090_Calc_SumWeight();
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

                    PS_PP090_FormItemEnabled();
                    PS_PP090_AddMatrixRow(oMat01.VisualRowCount, false);
                    //UDO방식
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP090H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP090L);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// CHOOSE_FROM_LIST 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
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
        /// EVENT_ROW_DELETE
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, SAPbouiCOM.IMenuEvent pVal, bool BubbleEvent)
        {
            int i;

            try
            {
                if (oLastColRow01 > 0)
                {
                    //Matrix 행삭제전 행삭제가능여부검사타기
                    if (pVal.BeforeAction == true)
                    {
                        if (PS_PP090_Validate("행삭제") == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        //Matrix 행삭제후 다시 행번채번등처리
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i; //행을 다시 순서대로정렬해서 행번에넣고(VisualCount값은 줄어든상태)
                        }
                        oMat01.FlushToDataSource(); //Matrix의 RowCount값 갯수도 줄어든 수만큼 갱신처리 해주며
                        oDS_PS_PP090L.RemoveRecord(oDS_PS_PP090L.Size - 1); //줄어든 행수만큼 DataSources값 갱신한 뒤
                        oMat01.LoadFromDataSource(); //그후 다시 데이터소스를 읽어와 화면완성을 한다.

                        //행이 없으면 한줄추가
                        if (oMat01.RowCount == 0)
                        {
                            PS_PP090_AddMatrixRow(0, false);
                        }
                        else
                        {
                            //현재행삭제한 행의PorNum값이 있는행지우면 넘어가고 없는 마지막행값지우면 한행추가
                            if (!string.IsNullOrEmpty(oDS_PS_PP090L.GetValue("U_LotNo", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_PP090_AddMatrixRow(oMat01.RowCount, false);
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
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            string errMessage = string.Empty;

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
                                if (PS_PP090_Validate("취소") == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }

                            }
                            else
                            {
                                errMessage = "현재 모드에서는 취소할수 없습니다.";
                                BubbleEvent = false;
                                throw new Exception();
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
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent);
                            PS_PP090_Calc_SumWeight();
                            break;
                        case "1281": //찾기
                            PS_PP090_FormItemEnabled();
                            break;
                        case "1282": //추가
                            PS_PP090_FormItemEnabled();
                            PS_PP090_AddMatrixRow(0, true);
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
                            PS_PP090_FormItemEnabled();
                            break;
                        case "1287": //복제
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
