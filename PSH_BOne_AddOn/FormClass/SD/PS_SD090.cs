using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 이동처리등록
	/// </summary>
	internal class PS_SD090 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01; 
		private SAPbouiCOM.DBDataSource oDS_PS_SD090H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_SD090L; //등록라인
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        
		// 입고 DI를 위한 정보를 가지는 구조체
		public class StockInfos
		{
			public string CardCode; //고객코드
			public string ItemCode; //품목코드
			public string FromWarehouseCode; //창고코드
			public string ToWarehouseCode; //창고코드
			public double Weight; //중량
			public double UnWeight;
			public string BatchNum; //배치번호
			public double BatchWeight;//배치중량
			public int Qty; //수량
			public string TransNo; //재고이전문서번호
			public string Chk;
			public int MatrixRow;
			public string StockTransDocEntry; //재고이전문서번호
			public string StockTransLineNum; //재고이전라인번호
			public string Indate; //전기일
        }

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD090.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD090_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD090");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);

				PS_SD090_CreateItems();
				PS_SD090_SetComboBox();
				PS_SD090_EnableFormItem();
				PS_SD090_SetDocEntry();
                PS_SD090_AddMatrixRow(0, true);

                oForm.EnableMenu("1293", true); //행삭제
				oForm.EnableMenu("1283", false); //제거
				oForm.EnableMenu("1284", false); //취소
				oForm.EnableMenu("1285", false); //복원
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
			}
		}

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_SD090_CreateItems()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oDS_PS_SD090H = oForm.DataSources.DBDataSources.Item("@PS_SD090H");
                oDS_PS_SD090L = oForm.DataSources.DBDataSources.Item("@PS_SD090L");
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.AutoResizeColumns();
                
                oForm.DataSources.UserDataSources.Add("SumQty", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("SumQty").Specific.DataBind.SetBound(true, "", "SumQty");

                oForm.DataSources.UserDataSources.Add("SumWeight", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("SumWeight").Specific.DataBind.SetBound(true, "", "SumWeight");

                oDS_PS_SD090H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));

                string userTempBPLID = dataHelpClass.User_BPLID();
                string userBPLID = (userTempBPLID == "1" || userTempBPLID == "4") ? userTempBPLID : "1";

                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL Where BPLId = '1' Or BPLId = '4' order by BPLId", userBPLID, false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_SD090_SetComboBox()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Set_ComboList(oForm.Items.Item("OutWhCd").Specific, "SELECT WhsCode, WhsName FROM [OWHS] order by WhsCode", "104", false, true); //출고창고
                dataHelpClass.Set_ComboList(oForm.Items.Item("InWhCd").Specific, "SELECT WhsCode, WhsName FROM [OWHS] order by WhsCode", "101", false, true); //입고창고

                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItmBsort"), "SELECT Code, Name FROM [@PSH_ITMBSORT] ORDER BY Code", "", ""); //품목대분류
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItmMsort"), "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] ORDER BY U_Code", "", ""); //품목중분류
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItemType"), "SELECT Code, Name FROM [@PSH_SHAPE] ORDER BY Code", "", ""); //형태타입
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("Quality"), "SELECT Code, Name FROM [@PSH_QUALITY] ORDER BY Code", "", ""); //질별
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("Mark"), "SELECT Code, Name FROM [@PSH_MARK] ORDER BY Code", "", ""); //인증기호
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("SbasUnit"), "SELECT Code, Name FROM [@PSH_UOMORG] ORDER BY Code", "", ""); //매입기준단위
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 각 모드에 따른 아이템설정
        /// </summary>
        private void PS_SD090_EnableFormItem()
        {
            string lQuery;
            SAPbobsCOM.Recordset lRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("RepName").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("ShipTo").Enabled = true;
                    oForm.Items.Item("CarCo").Enabled = true;
                    oForm.Items.Item("CarNo").Enabled = true;
                    oForm.Items.Item("ArrSite").Enabled = true;
                    oForm.Items.Item("Fare").Enabled = true;
                    oForm.Items.Item("Specific").Enabled = true;
                    oForm.Items.Item("ChulPrin").Enabled = false;
                    oMat01.Columns.Item("ItemCode").Editable = true;
                    oMat01.Columns.Item("ItemGu").Editable = true;
                    oForm.Items.Item("DocDate").Specific.String = DateTime.Now.ToString("yyyyMMdd");
                    oForm.Items.Item("RepName").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("RepName").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("ShipTo").Enabled = true;
                    oForm.Items.Item("CarCo").Enabled = true;
                    oForm.Items.Item("CarNo").Enabled = true;
                    oForm.Items.Item("ArrSite").Enabled = true;
                    oForm.Items.Item("Fare").Enabled = false;
                    oForm.Items.Item("Specific").Enabled = false;
                    oForm.Items.Item("ChulPrin").Enabled = true;
                    oMat01.Columns.Item("ItemCode").Editable = false;
                    oMat01.Columns.Item("ItemGu").Editable = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("CardCode").Enabled = false;
                    oForm.Items.Item("BPLId").Enabled = false;
                    oForm.Items.Item("RepName").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = false;
                    oForm.Items.Item("ShipTo").Enabled = false;
                    oForm.Items.Item("CarCo").Enabled = false;
                    oForm.Items.Item("CarNo").Enabled = false;
                    oForm.Items.Item("ArrSite").Enabled = false;
                    oForm.Items.Item("Fare").Enabled = false;
                    oForm.Items.Item("Specific").Enabled = false;
                    oForm.Items.Item("ChulPrin").Enabled = true;
                    oMat01.Columns.Item("ItemCode").Editable = false;
                    oMat01.Columns.Item("ItemGu").Editable = false;

                    lQuery = "SELECT Status,Canceled FROM [@PS_SD090H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                    lRecordSet.DoQuery(lQuery);
                    if (lRecordSet.Fields.Item(0).Value == "O")
                    {
                        oForm.Items.Item("Status").Specific.Value = "미결";
                    }
                    else if (lRecordSet.Fields.Item(0).Value == "C")
                    {
                        if (lRecordSet.Fields.Item(1).Value == "Y")
                        {
                            oForm.Items.Item("Status").Specific.Value = "취소";
                        }
                        else
                        {
                            oForm.Items.Item("Status").Specific.Value = "종료";
                        }
                    }
                }
                oMat01.AutoResizeColumns();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(lRecordSet);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_SD090_SetDocEntry()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD090'", "");
                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
                {
                    oForm.Items.Item("DocEntry").Specific.Value = "1";
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 메트릭스 Row추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_SD090_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                if (RowIserted == false)
                {
                    oDS_PS_SD090L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_SD090L.Offset = oRow;
                oDS_PS_SD090L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
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
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_SD090_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            try
            {
                switch (oUID)
                {
                    case "Mat01":
                        if ((oRow == oMat01.RowCount || oMat01.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim()))
                        {
                            oMat01.FlushToDataSource();
                            PS_SD090_AddMatrixRow(oMat01.RowCount, false);
                            oMat01.Columns.Item("ItemCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        break;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 필수입력사항 체크(헤더)
        /// </summary>
        /// <returns></returns>
        private bool PS_SD090_DelHeaderSpaceLine()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_SD090H.GetValue("U_DocDate", 0))) 
                {
                    errMessage = "거래일자는 필수입력 사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_SD090H.GetValue("U_BPLId", 0)))
                {
                    errMessage = "사업장은 필수입력 사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_SD090H.GetValue("U_OutWhCd", 0)))
                {
                    errMessage = "출고창고는 필수입력 사항입니다. 확인하세요.";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oDS_PS_SD090H.GetValue("U_InWhCd", 0)))
                {
                    errMessage = "입고창고는 필수입력 사항입니다. 확인하세요.";
                    throw new Exception();
                }

                returnValue = true;
            }
            catch(Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// 필수입력사항 체크(라인)
        /// </summary>
        /// <returns></returns>
        private bool PS_SD090_DelMatrixSpaceLine()
        {
            bool returnValue = false;
            int i;
            string errMessage = string.Empty;

            try
            {
                oMat01.FlushToDataSource();

                if (oMat01.VisualRowCount <= 1)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                if (oMat01.VisualRowCount > 0)
                {
                    for (i = 0; i <= oMat01.VisualRowCount - 2; i++)
                    {
                        oDS_PS_SD090L.Offset = i;
                        if (string.IsNullOrEmpty(oDS_PS_SD090L.GetValue("U_ItemCode", i)))
                        {
                            errMessage = "아이템 데이터는 필수입니다. 확인하세요.";
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PS_SD090L.GetValue("U_Qty", i)))
                        {
                            errMessage = "수량은 필수입니다. 확인하세요.";
                            throw new Exception();
                        }
                        else if(string.IsNullOrEmpty(oDS_PS_SD090L.GetValue("U_SD091HNo", i)))
                        {
                            errMessage = "이동요청 문서는 필수입니다. 확인하세요.";
                            throw new Exception();
                        }
                    }

                    if (string.IsNullOrEmpty(oDS_PS_SD090L.GetValue("U_SD091HNo", oMat01.VisualRowCount - 1)))
                    {
                        oDS_PS_SD090L.RemoveRecord(oMat01.VisualRowCount - 1);
                    }
                }
                
                oMat01.LoadFromDataSource();

                returnValue = true;
            }
            catch(Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// 재고이전
        /// </summary>
        /// <returns></returns>
        private bool PS_SD090_TransStock()
        {
            bool returnValue = false;
            string errCode = string.Empty;
            string errDIMsg = string.Empty;
            int errDICode = 0;
            int RetVal;
            int StockTransLineCounter;
            int Q;
            int i;
            int K;
            int r;
            int DocCnt = 0;
            string Chk1_Val;
            string sCur_ItemCode;
            string sNxt_ItemCode;
            string sCur_TrCardCode;
            string sCur_TrOutWhs;
            string sNxt_TrOutWhs;
            string sCur_TrInWhs;
            string sNxt_TrInWhs;

            SAPbobsCOM.StockTransfer oStockTrans = null;
            SAPbouiCOM.ProgressBar oPrgBar = null;
            
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oPrgBar = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                PSH_Globals.oCompany.StartTransaction();

                //현재월의 전기기간 체크 후 잠겨있으면 DI API 미실행
                if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
                {
                    errCode = "2";
                    throw new Exception();
                }

                List<StockInfos> stockInfoList = new List<StockInfos>(); //재고

                for (i = 0; i <= oMat01.RowCount - 1; i++)
                {
                    StockInfos stockInfo = new StockInfos
                    {
                        CardCode = oDS_PS_SD090H.GetValue("U_CardCode", 0).ToString().Trim(),
                        ItemCode = oDS_PS_SD090L.GetValue("U_ItemCode", i).ToString().Trim(),
                        FromWarehouseCode = oDS_PS_SD090H.GetValue("U_OutWhCd", 0).ToString().Trim(),
                        ToWarehouseCode = oDS_PS_SD090H.GetValue("U_InWhCd", 0).ToString().Trim(),
                        BatchNum = oDS_PS_SD090L.GetValue("U_BatchNum", i).ToString().Trim(),
                        Weight = System.Math.Round(Convert.ToDouble(oDS_PS_SD090L.GetValue("U_Weight", i).ToString().Trim()), 2),
                        UnWeight = System.Math.Round(Convert.ToDouble(oDS_PS_SD090L.GetValue("U_Unweight", i).ToString().Trim()), 2),
                        BatchWeight = System.Math.Round(Convert.ToDouble(oDS_PS_SD090L.GetValue("U_Qty", i).ToString().Trim()), 2),
                        Qty = Convert.ToInt32(oDS_PS_SD090L.GetValue("U_Qty", i).ToString().Trim()),
                        TransNo = oForm.Items.Item("DocEntry").Specific.Value + (i + 1),
                        Chk = "N",
                        MatrixRow = i + 1,
                        Indate = oForm.Items.Item("DocDate").Specific.Value
                    };

                    stockInfoList.Add(stockInfo);
                }

                for (i = 0; i < stockInfoList.Count; i++)
                {
                    stockInfoList[i].StockTransDocEntry = "";
                }

                for (i = 0; i < stockInfoList.Count; i++)
                {
                    Chk1_Val = stockInfoList[i].Chk;

                    if (Chk1_Val != "N")
                    {
                        continue;
                    }

                    sCur_TrOutWhs = stockInfoList[i].FromWarehouseCode;

                    oStockTrans = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
                    oStockTrans.CardCode = stockInfoList[i].CardCode;
                    oStockTrans.DocDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(stockInfoList[i].Indate, "-"));
                    oStockTrans.FromWarehouse = sCur_TrOutWhs;
                    oStockTrans.Comments = "재고이전" + oForm.Items.Item("DocEntry").Specific.Value + ".";

                    StockTransLineCounter = -1;

                    for (K = i; K < stockInfoList.Count; K++)
                    {
                        Chk1_Val = stockInfoList[K].Chk;

                        if (Chk1_Val != "N")
                        {
                            continue;
                        }
                        sCur_TrCardCode = stockInfoList[K].CardCode;
                        sNxt_TrOutWhs = stockInfoList[K].FromWarehouseCode;
                        sCur_ItemCode = stockInfoList[K].ItemCode;
                        sCur_TrInWhs = stockInfoList[K].ToWarehouseCode;

                        if (sCur_TrOutWhs != sNxt_TrOutWhs)
                        {
                            continue;
                        }

                        if (i != K)
                        {
                            oStockTrans.Lines.Add();
                        }
                        StockTransLineCounter += 1;

                        oStockTrans.Lines.ItemCode = stockInfoList[K].ItemCode;
                        oStockTrans.Lines.UserFields.Fields.Item("U_Qty").Value = stockInfoList[K].Qty;
                        oStockTrans.Lines.UserFields.Fields.Item("U_UnWeight").Value = stockInfoList[K].UnWeight;
                        oStockTrans.Lines.Quantity = stockInfoList[K].Qty;
                        oStockTrans.Lines.WarehouseCode = stockInfoList[K].ToWarehouseCode;
                        oStockTrans.Lines.UserFields.Fields.Item("U_BatchNum").Value = stockInfoList[K].BatchNum;
                        oStockTrans.Lines.BatchNumbers.BatchNumber = stockInfoList[K].BatchNum;
                        oStockTrans.Lines.BatchNumbers.Quantity = System.Math.Round(stockInfoList[K].BatchWeight, 2);
                        oStockTrans.Lines.BatchNumbers.Notes = "재고이전(Addon)";
                        stockInfoList[K].Chk = "Y";
                        stockInfoList[K].StockTransDocEntry = "Checked";
                        stockInfoList[K].StockTransLineNum = Convert.ToString(StockTransLineCounter);

                        for (Q = K + 1; Q < stockInfoList.Count; Q++)
                        {
                            Chk1_Val = stockInfoList[Q].Chk;

                            if (Chk1_Val != "N")
                            {
                                continue;
                            }

                            sNxt_TrOutWhs = stockInfoList[Q].FromWarehouseCode;
                            sNxt_ItemCode = stockInfoList[Q].ItemCode;
                            sNxt_TrInWhs = stockInfoList[Q].ToWarehouseCode;

                            if (sNxt_TrOutWhs == sCur_TrOutWhs & sCur_ItemCode == sNxt_ItemCode & sCur_TrInWhs == sNxt_TrInWhs)
                            {
                                if (dataHelpClass.GetValue("SELECT ManBatchNum FROM OITM WHERE ITEMCODE = ''", 0, 1) == "Y")
                                {
                                    oStockTrans.Lines.BatchNumbers.Add();
                                    oStockTrans.Lines.BatchNumbers.BatchNumber = stockInfoList[Q].BatchNum;
                                    oStockTrans.Lines.BatchNumbers.Quantity = System.Math.Round(stockInfoList[Q].BatchWeight, 2);
                                    oStockTrans.Lines.UserFields.Fields.Item("Quantity").Value = Convert.ToInt32(oStockTrans.Lines.UserFields.Fields.Item("Quantity").Value) + stockInfoList[Q].Qty;
                                    oStockTrans.Lines.Quantity = oStockTrans.Lines.Quantity + System.Math.Round(stockInfoList[Q].Weight, 2);
                                    oStockTrans.Lines.BatchNumbers.Notes = "재고이전(Addon)";
                                    stockInfoList[Q].Chk = "Y";
                                    stockInfoList[Q].StockTransDocEntry = "Checked";
                                    stockInfoList[Q].StockTransLineNum = Convert.ToString(StockTransLineCounter);
                                }
                            }
                        }
                    }

                    RetVal = oStockTrans.Add();

                    if (RetVal == 0)
                    {
                        DocCnt += 1;
                        PSH_Globals.oCompany.GetNewObjectCode(out string afterDIDocNum);

                        for (r = 0; r < stockInfoList.Count; r++)
                        {
                            if (stockInfoList[r].StockTransDocEntry == "Checked")
                            {
                                stockInfoList[r].StockTransDocEntry = afterDIDocNum;
                                oDS_PS_SD090H.SetValue("U_StoTrDoc", 0, stockInfoList[r].StockTransDocEntry);
                            }
                        }
                    }
                    else
                    {
                        PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                        errCode = "1";
                        throw new Exception();
                    }
                }

                PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                PSH_Globals.SBO_Application.StatusBar.SetText(DocCnt + " 개의 재고이전 문서가 발행되었습니다 !", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

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
                    PSH_Globals.SBO_Application.MessageBox("DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.MessageBox("현재월의 전기기간이 잠겼습니다. 회계부서에 문의하세요.");
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                if (oStockTrans != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oStockTrans);
                }

                if (oPrgBar != null)
                {
                    oPrgBar.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oPrgBar);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// 리포트 출력
        /// </summary>
        [STAThread]
        private void PS_SD090_Print_Report01()
        {
            string WinTitle;
            string ReportName;
            string docEntry;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                docEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

                WinTitle = "[PS_SD090] 출고원부/반출증";
                ReportName = "PS_SD090_10.rpt";
                //프로시저 : PS_SD090_10

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>
                {
                    new PSH_DataPackClass("@DocEntry", docEntry)
                };

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
            }
            finally
            {
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
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    //Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_SD090_DelHeaderSpaceLine() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_SD090_DelMatrixSpaceLine() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) //재고 이동 DI API
                            {
                                if (PS_SD090_TransStock() == false)
                                {
                                    PS_SD090_AddMatrixRow(oMat01.VisualRowCount, true);
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                        }
                    }
                    else if (pVal.ItemUID == "ChulPrin")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PS_SD090_Print_Report01);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                    else
                    {
                        if (pVal.ItemChanged == true)
                        {
                            if (pVal.ItemUID == "Mat01" && pVal.ColUID == "ItemCode")
                            {
                                PS_SD090_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                            }
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
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == false)
                        {
                            PS_SD090_EnableFormItem();
                            PS_SD090_AddMatrixRow(0, true);
                        }
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
                    //거래처코드
                    if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value))
                    {
                        if (pVal.ItemUID == "CardCode" && pVal.CharPressed == 9)
                        {
                            oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }

                    //아이템코드
                    if (pVal.ItemUID == "Mat01" && pVal.ColUID == "SD091HNo" && pVal.CharPressed == 9)
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String))
                        {
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }

                    //담당자
                    if (string.IsNullOrEmpty(oForm.Items.Item("RepName").Specific.Value))
                    {
                        if (pVal.ItemUID == "RepName" && pVal.CharPressed == 9)
                        {
                            oForm.Items.Item("RepName").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }

                    //납품처
                    if (string.IsNullOrEmpty(oForm.Items.Item("ShipTo").Specific.Value))
                    {
                        if (pVal.ItemUID == "ShipTo" && pVal.CharPressed == 9)
                        {
                            oForm.Items.Item("ShipTo").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }

                    //출고창고
                    if (pVal.ItemUID == "OutWhCd" && pVal.CharPressed == 9)
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.String))
                        {
                            oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }

                    //입고창고
                    if (pVal.ItemUID == "InWhCd" && pVal.CharPressed == 9)
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.String))
                        {
                            oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "BPLId")
                        {
                            if (oForm.Items.Item("BPLId").Specific.Value == "1")
                            {
                                oDS_PS_SD090H.SetValue("U_OutWhCd", 0, "101");
                                oDS_PS_SD090H.SetValue("U_InWhCd", 0, "104");
                            }
                            else if (oForm.Items.Item("BPLId").Specific.Value == "4")
                            {
                                oDS_PS_SD090H.SetValue("U_OutWhCd", 0, "104");
                                oDS_PS_SD090H.SetValue("U_InWhCd", 0, "101");
                            }
                        }
                        oForm.Update();
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
        /// DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01" && pVal.ColUID == "ItemCode")
                        {
                            PS_SD090_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry;
            int docEntry;
            int lineID;
            double SumQty = 0;
            double SumWeight = 0;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        //거래처 이름 Query
                        if (pVal.ItemUID == "CardCode")
                        {
                            sQry = "Select CardName From [OCRD] Where CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
                            oRecordSet.DoQuery(sQry);
                            oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                        }

                        //사원 이름 Query
                        if (pVal.ItemUID == "RepName")
                        {
                            sQry = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item("RepName").Specific.Value.ToString().Trim() + "'";
                            oRecordSet.DoQuery(sQry);
                            oForm.Items.Item("RepNm1").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                        }

                        //납품처 이름 Query
                        if (pVal.ItemUID == "ShipTo")
                        {
                            sQry = "Select CardName From [OCRD] Where CardCode = '" + oForm.Items.Item("ShipTo").Specific.Value.ToString().Trim() + "'";
                            oRecordSet.DoQuery(sQry);
                            oForm.Items.Item("ShipNm").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                        }

                        //아이템 코드
                        if (pVal.ItemUID == "Mat01" && pVal.ColUID == "ItemCode")
                        {
                            PS_SD090_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                        }

                        //출고 창고
                        if (pVal.ItemUID == "OutWhCd")
                        {
                            sQry = "Select WhsName From [OWHS] Where WhsCode = '" + oForm.Items.Item("OutWhCd").Specific.Value.ToString().Trim() + "'";
                            oRecordSet.DoQuery(sQry);
                            oForm.Items.Item("OutWhName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                        }

                        //입고 창고
                        if (pVal.ItemUID == "InWhCd")
                        {
                            sQry = "Select WhsName From [OWHS] Where WhsCode = '" + oForm.Items.Item("InWhCd").Specific.Value.ToString().Trim() + "'";
                            oRecordSet.DoQuery(sQry);
                            oForm.Items.Item("InWhCd").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                        }

                        //Matrix-작업요청
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "SD091HNo")
                            {
                                docEntry = Convert.ToInt32(oMat01.Columns.Item("SD091HNo").Cells.Item(pVal.Row).Specific.Value.ToString().Split("-")[0]);
                                lineID = Convert.ToInt32(oMat01.Columns.Item("SD091HNo").Cells.Item(pVal.Row).Specific.Value.ToString().Split("-")[1]);

                                sQry = "  SELECT    U_ItemCode,";
                                sQry += "           U_ItemName,";
                                sQry += "           U_ItemGu,";
                                sQry += "           U_Qty,";
                                sQry += "           U_Unweight,";
                                sQry += "           U_Weight,";
                                sQry += "           U_Comments,";
                                sQry += "           U_ItmBsort,";
                                sQry += "           U_ItmMsort,";
                                sQry += "           U_Unit1,";
                                sQry += "           U_Size,";
                                sQry += "           U_ItemType,";
                                sQry += "           U_Quality,";
                                sQry += "           U_Mark,";
                                sQry += "           U_CallSize,";
                                sQry += "           U_SbasUnit ";
                                sQry += " FROM      [@PS_SD091L]";
                                sQry += " WHERE     DocEntry = '" + docEntry + "'";
                                sQry += "           AND LineId = '" + lineID + "'";
                                oRecordSet.DoQuery(sQry);
                                oMat01.FlushToDataSource();
                                oDS_PS_SD090L.SetValue("U_ItemCode", pVal.Row - 1, oRecordSet.Fields.Item(0).Value.ToString().Trim());
                                oDS_PS_SD090L.SetValue("U_ItemName", pVal.Row - 1, oRecordSet.Fields.Item(1).Value.ToString().Trim());
                                oDS_PS_SD090L.SetValue("U_ItemGu", pVal.Row - 1, oRecordSet.Fields.Item(2).Value.ToString().Trim());
                                oDS_PS_SD090L.SetValue("U_Qty", pVal.Row - 1, oRecordSet.Fields.Item(3).Value.ToString().Trim());
                                oDS_PS_SD090L.SetValue("U_Unweight", pVal.Row - 1, oRecordSet.Fields.Item(4).Value.ToString().Trim());
                                oDS_PS_SD090L.SetValue("U_Weight", pVal.Row - 1, oRecordSet.Fields.Item(5).Value.ToString().Trim());
                                oDS_PS_SD090L.SetValue("U_Comments", pVal.Row - 1, oRecordSet.Fields.Item(6).Value.ToString().Trim());
                                oDS_PS_SD090L.SetValue("U_ItmBsort", pVal.Row - 1, oRecordSet.Fields.Item(7).Value.ToString().Trim());
                                oDS_PS_SD090L.SetValue("U_ItmMsort", pVal.Row - 1, oRecordSet.Fields.Item(8).Value.ToString().Trim());
                                oDS_PS_SD090L.SetValue("U_Unit1", pVal.Row - 1, oRecordSet.Fields.Item(9).Value.ToString().Trim());
                                oDS_PS_SD090L.SetValue("U_Size", pVal.Row - 1, oRecordSet.Fields.Item(10).Value.ToString().Trim());
                                oDS_PS_SD090L.SetValue("U_ItemType", pVal.Row - 1, oRecordSet.Fields.Item(11).Value.ToString().Trim());
                                oDS_PS_SD090L.SetValue("U_Quality", pVal.Row - 1, oRecordSet.Fields.Item(12).Value.ToString().Trim());
                                oDS_PS_SD090L.SetValue("U_Mark", pVal.Row - 1, oRecordSet.Fields.Item(13).Value.ToString().Trim());
                                oDS_PS_SD090L.SetValue("U_CallSize", pVal.Row - 1, oRecordSet.Fields.Item(14).Value.ToString().Trim());
                                oDS_PS_SD090L.SetValue("U_SbasUnit", pVal.Row - 1, oRecordSet.Fields.Item(15).Value.ToString().Trim());

                                PS_SD090_AddMatrixRow(oMat01.VisualRowCount, false);
                                for (int i = 1; i <= oMat01.VisualRowCount; i++)
                                {
                                    if (string.IsNullOrEmpty(oDS_PS_SD090L.GetValue("U_SD091HNo", i - 1)))
                                    {
                                        oDS_PS_SD090L.RemoveRecord(i - 1);
                                        oMat01.LoadFromDataSource();
                                    }
                                }

                                sQry = "Select U_OutWhCd, U_InWhCd From [@PS_SD091H] Where DocEntry = '" + docEntry + "'";
                                oRecordSet.DoQuery(sQry);
                                oDS_PS_SD090H.SetValue("U_OutWhCd", pVal.Row - 1, oRecordSet.Fields.Item(0).Value.ToString().Trim());
                                oDS_PS_SD090H.SetValue("U_InWhCd", pVal.Row - 1, oRecordSet.Fields.Item(1).Value.ToString().Trim());

                                for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    if (!string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value))
                                    {
                                        SumQty += Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value);
                                    }
                                    SumWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value);
                                }
                                oForm.Items.Item("SumQty").Specific.Value = SumQty;
                                oForm.Items.Item("SumWeight").Specific.Value = SumWeight;

                                PS_SD090_AddMatrixRow(oMat01.VisualRowCount, false);
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
            double SumQty = 0;
            double SumWeight = 0;

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        if (!string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value))
                        {
                            SumQty += Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value);
                        }

                        SumWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value);
                    }

                    oForm.Items.Item("SumQty").Specific.Value = SumQty;
                    oForm.Items.Item("SumWeight").Specific.Value = SumWeight;

                    PS_SD090_EnableFormItem();
                    PS_SD090_AddMatrixRow(oMat01.VisualRowCount, false);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD090H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD090L);
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
        /// 행삭제 체크 메소드(Raise_FormMenuEvent 에서 사용)
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            int i;
            double SumQty = 0;
            double SumWeight = 0;

            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }
                        oMat01.FlushToDataSource();
                        oDS_PS_SD090L.RemoveRecord(oDS_PS_SD090L.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PS_SD090_AddMatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_SD090L.GetValue("U_ItemCode", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_SD090_AddMatrixRow(oMat01.RowCount, false);
                            }
                        }

                        for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                        {
                            if (!string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value))
                            {
                                SumQty += Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value);
                            }
                            SumWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value);

                        }
                        oForm.Items.Item("SumQty").Specific.Value = SumQty;
                        oForm.Items.Item("SumWeight").Specific.Value = SumWeight;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            double SumQty = 0;
            double SumWeight = 0;

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
                        case "1281": //찾기
                            PS_SD090_EnableFormItem();
                            break;
                        case "1282": //추가
                            PS_SD090_EnableFormItem();
                            PS_SD090_AddMatrixRow(0, true);
                            PS_SD090_SetDocEntry();
                            oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            //PS_SD090_EnableFormItem();
                            //if (oMat01.VisualRowCount > 0)
                            //{
                            //    if (!string.IsNullOrEmpty(oMat01.Columns.Item("SD091HNo").Cells.Item(oMat01.VisualRowCount).Specific.Value))
                            //    {
                            //        PS_SD090_AddMatrixRow(oMat01.RowCount, true);
                            //    }
                            //}
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
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
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                        //Raise_EVENT_FORM_DATA_LOAD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                        //Raise_EVENT_FORM_DATA_ADD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                        //Raise_EVENT_FORM_DATA_UPDATE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                        //Raise_EVENT_FORM_DATA_DELETE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
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
