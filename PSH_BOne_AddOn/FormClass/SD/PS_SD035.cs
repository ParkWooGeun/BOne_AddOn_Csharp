using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using System.Collections.Generic;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 출하리스트 등록 및 조회
	/// </summary>
	internal class PS_SD035 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_SD035H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_SD035L; //등록라인
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD035.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD035_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD035");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				//oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);

				PS_SD035_CreateItems();
				PS_SD035_SetComboBox();
				PS_SD035_SetInitial();
				PS_SD035_EnableMenus();
				PS_SD035_SetDocument(oFormDocEntry);
				//PS_SD035_ResizeForm();

				oForm.EnableMenu("1283", false); //삭제
				oForm.EnableMenu("1287", false); //복제
				oForm.EnableMenu("1286", false); //닫기
				oForm.EnableMenu("1284", false); //취소
				oForm.EnableMenu("1293", false); //행삭제
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
        private void PS_SD035_CreateItems()
        {
            try
            {
                oDS_PS_SD035L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oMat01 = oForm.Items.Item("Mat01").Specific;

                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                ////////////대상조회//////////_S
                //수주일자(시작)
                oForm.DataSources.UserDataSources.Add("ORDR_FrDt", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("ORDR_FrDt").Specific.DataBind.SetBound(true, "", "ORDR_FrDt");

                //수주일자(종료)
                oForm.DataSources.UserDataSources.Add("ORDR_ToDt", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("ORDR_ToDt").Specific.DataBind.SetBound(true, "", "ORDR_ToDt");

                //생산완료일자(시작)
                oForm.DataSources.UserDataSources.Add("PP080_FrDt", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("PP080_FrDt").Specific.DataBind.SetBound(true, "", "PP080_FrDt");

                //생산완료일자(종료)
                oForm.DataSources.UserDataSources.Add("PP080_ToDt", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("PP080_ToDt").Specific.DataBind.SetBound(true, "", "PP080_ToDt");

                //거래처구분
                oForm.DataSources.UserDataSources.Add("CardType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CardType").Specific.DataBind.SetBound(true, "", "CardType");

                //품목구분
                oForm.DataSources.UserDataSources.Add("ItemType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ItemType").Specific.DataBind.SetBound(true, "", "ItemType");

                //출고구분(수기,정상)
                oForm.DataSources.UserDataSources.Add("OutCls", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("OutCls").Specific.DataBind.SetBound(true, "", "OutCls");

                //수주처
                oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("CardCode").Specific.DataBind.SetBound(true, "", "CardCode");

                //수주처명
                oForm.DataSources.UserDataSources.Add("CardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("CardName").Specific.DataBind.SetBound(true, "", "CardName");

                //품목코드(작번)
                oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

                //품목명
                oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");
                ////////////대상조회//////////_E

                ////////////출하리스트조회//////////_S
                //출하일자(시작)
                oForm.DataSources.UserDataSources.Add("Out_FrDt", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("Out_FrDt").Specific.DataBind.SetBound(true, "", "Out_FrDt");

                //출하일자(종료)
                oForm.DataSources.UserDataSources.Add("Out_ToDt", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("Out_ToDt").Specific.DataBind.SetBound(true, "", "Out_ToDt");

                //거래처구분
                oForm.DataSources.UserDataSources.Add("SCardType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SCardType").Specific.DataBind.SetBound(true, "", "SCardType");

                //품목구분
                oForm.DataSources.UserDataSources.Add("SItemType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SItemType").Specific.DataBind.SetBound(true, "", "SItemType");

                //출고구분(수기,정상)
                oForm.DataSources.UserDataSources.Add("SOutCls", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SOutCls").Specific.DataBind.SetBound(true, "", "SOutCls");

                //수주처
                oForm.DataSources.UserDataSources.Add("SCardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("SCardCode").Specific.DataBind.SetBound(true, "", "SCardCode");

                //수주처명
                oForm.DataSources.UserDataSources.Add("SCardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("SCardName").Specific.DataBind.SetBound(true, "", "SCardName");

                //품목코드(작번)
                oForm.DataSources.UserDataSources.Add("SItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("SItemCode").Specific.DataBind.SetBound(true, "", "SItemCode");

                //품목명
                oForm.DataSources.UserDataSources.Add("SItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("SItemName").Specific.DataBind.SetBound(true, "", "SItemName");
                ////////////출하리스트조회//////////_E
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_SD035_SetComboBox()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                //거래처구분(대상조회)
                sQry = "  SELECT    U_Minor,";
                sQry += "           U_CdName";
                sQry += " FROM      [@PS_SY001L]";
                sQry += " WHERE     Code = 'C100'";
                sQry += "           AND U_UseYN = 'Y'";
                sQry += " ORDER BY  U_Seq";

                oForm.Items.Item("CardType").Specific.ValidValues.Add("%", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("CardType").Specific, sQry, "", false, false);
                oForm.Items.Item("CardType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //거래처구분(출하리스트조회)
                oForm.Items.Item("SCardType").Specific.ValidValues.Add("%", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("SCardType").Specific, sQry, "", false, false);
                oForm.Items.Item("SCardType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //품목구분(대상조회)
                sQry = "  SELECT    U_Minor,";
                sQry += "           U_CdName";
                sQry += " FROM      [@PS_SY001L]";
                sQry += " WHERE     Code = 'S002'";
                sQry += "           AND U_UseYN = 'Y'";
                sQry += " ORDER BY  U_Seq";

                oForm.Items.Item("ItemType").Specific.ValidValues.Add("%", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType").Specific, sQry, "", false, false);
                oForm.Items.Item("ItemType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //품목구분(출하리스트조회)
                oForm.Items.Item("SItemType").Specific.ValidValues.Add("%", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("SItemType").Specific, sQry, "", false, false);
                oForm.Items.Item("SItemType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //출고구분(대상조회)
                oForm.Items.Item("OutCls").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("OutCls").Specific.ValidValues.Add("Y", "정상출고");
                oForm.Items.Item("OutCls").Specific.ValidValues.Add("N", "수기출고");
                oForm.Items.Item("OutCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //출고구분(출하리스트조회)
                oForm.Items.Item("SOutCls").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("SOutCls").Specific.ValidValues.Add("Y", "정상출고");
                oForm.Items.Item("SOutCls").Specific.ValidValues.Add("N", "수기출고");
                oForm.Items.Item("SOutCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
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
        /// 화면 초기화
        /// </summary>
        private void PS_SD035_SetInitial()
        {
            try
            {
                //출하일자
                oForm.Items.Item("Out_FrDt").Specific.Value = DateTime.Now.ToString("yyyyMMdd"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD");
                oForm.Items.Item("Out_ToDt").Specific.Value = DateTime.Now.ToString("yyyyMMdd"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 메뉴활성화
        /// </summary>
        private void PS_SD035_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, false, false, false, false, false, false, false);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        private void PS_SD035_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_SD035_EnableFormItem();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 각모드에따른 아이템설정
        /// </summary>
        private void PS_SD035_EnableFormItem()
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("Mat01").Enabled = true;
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocEntry").Specific.Value = "";
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = false;
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = true;
                }

                oMat01.AutoResizeColumns();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메트릭스 Row추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_SD035_AddMatrixRow(int oRow, ref bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_SD035L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_SD035L.Offset = oRow;
                oDS_PS_SD035L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
                oForm.Freeze(false);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 필수입력사항 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_SD035_CheckDataValid()
        {
            bool returnValue = false;
            string errMessage = string.Empty;            

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("ChkWkCd").Specific.Value))
                {
                    errMessage = "검사자가 입력되지 않았습니다.";
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
                {
                    errMessage = "등록일자가 입력되지 않았습니다.";
                    throw new Exception();
                }

                //라인정보 미입력 시
                if (oMat01.VisualRowCount == 1)
                {
                    errMessage = "라인이 존재하지 않습니다.";
                    throw new Exception();
                }

                //for (int i = 1; i <= oMat01.VisualRowCount - 1; i++)
                //{
                //}

                oMat01.FlushToDataSource();
                oDS_PS_SD035L.RemoveRecord(oDS_PS_SD035L.Size - 1);
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
        /// 데이터 조회
        /// </summary>
        private void PS_SD035_MTX01()
        {
            int i;
            string Query01;
            string ORDR_FrDt; //수주일자(시작)
            string ORDR_ToDt; //수주일자(종료)
            string PP080_FrDt; //생산완료일자(시작)
            string PP080_ToDt; //생산완료일자(종료)
            string CardType; //거래처구분
            string ItemType; //품목구분
            string OutCls; //출고구분
            string CardCode; //수주처
            string ItemCode; //작번
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                ORDR_FrDt = oForm.Items.Item("ORDR_FrDt").Specific.Value.ToString().Trim(); //수주일자(시작)
                ORDR_ToDt = oForm.Items.Item("ORDR_ToDt").Specific.Value.ToString().Trim(); //수주일자(종료)
                PP080_FrDt = oForm.Items.Item("PP080_FrDt").Specific.Value.ToString().Trim(); //생산완료일자(시작)
                PP080_ToDt = oForm.Items.Item("PP080_ToDt").Specific.Value.ToString().Trim(); //생산완료일자(종료)
                CardType = oForm.Items.Item("CardType").Specific.Value.ToString().Trim(); //거래처구분
                ItemType = oForm.Items.Item("ItemType").Specific.Value.ToString().Trim(); //품목구분
                OutCls = oForm.Items.Item("OutCls").Specific.Value.ToString().Trim(); //출고구분
                CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim(); //수주처
                ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim(); //수주처

                Query01 = "EXEC PS_SD035_01 '";
                Query01 += ORDR_FrDt + "','";
                Query01 += ORDR_ToDt + "','";
                Query01 += PP080_FrDt + "','";
                Query01 += PP080_ToDt + "','";
                Query01 += CardType + "','";
                Query01 += ItemType + "','";
                Query01 += OutCls + "','";
                Query01 += CardCode + "','";
                Query01 += ItemCode + "'";

                RecordSet01.DoQuery(Query01);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (RecordSet01.RecordCount == 0)
                {
                    oMat01.Clear();
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }

                for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_SD035L.InsertRecord(i);
                    }
                    oDS_PS_SD035L.Offset = i;
                    oDS_PS_SD035L.SetValue("U_LineNum", i, Convert.ToString(i + 1)); //라인번호
                    oDS_PS_SD035L.SetValue("U_ColReg01", i, RecordSet01.Fields.Item("Check").Value); //선택
                    oDS_PS_SD035L.SetValue("U_ColReg02", i, RecordSet01.Fields.Item("ItemCode").Value); //작번(품목코드)
                    oDS_PS_SD035L.SetValue("U_ColReg03", i, RecordSet01.Fields.Item("ItemName").Value); //품명
                    oDS_PS_SD035L.SetValue("U_ColReg04", i, RecordSet01.Fields.Item("CardType").Value); //거래처구분
                    oDS_PS_SD035L.SetValue("U_ColReg05", i, RecordSet01.Fields.Item("ItemType").Value); //품목구분
                    oDS_PS_SD035L.SetValue("U_ColReg06", i, RecordSet01.Fields.Item("CardCode").Value); //수주처코드
                    oDS_PS_SD035L.SetValue("U_ColReg07", i, RecordSet01.Fields.Item("CardName").Value); //수주처
                    oDS_PS_SD035L.SetValue("U_ColReg18", i, RecordSet01.Fields.Item("ORDR_LotNo").Value); //수주주문번호
                    oDS_PS_SD035L.SetValue("U_ColReg08", i, RecordSet01.Fields.Item("ORDR_Date").Value); //수주일자
                    oDS_PS_SD035L.SetValue("U_ColQty01", i, RecordSet01.Fields.Item("ORDR_Qty").Value); //수주수량
                    oDS_PS_SD035L.SetValue("U_ColReg10", i, RecordSet01.Fields.Item("PP080_Date").Value); //생산완료일자
                    oDS_PS_SD035L.SetValue("U_ColQty02", i, RecordSet01.Fields.Item("PP080_Qty").Value); //생산완료수량
                    oDS_PS_SD035L.SetValue("U_ColReg12", i, RecordSet01.Fields.Item("PP080_YN").Value); //생산완료여부
                    oDS_PS_SD035L.SetValue("U_ColReg13", i, RecordSet01.Fields.Item("QM4XX_Date").Value); //검사일자(최종)
                    oDS_PS_SD035L.SetValue("U_ColQty03", i, RecordSet01.Fields.Item("QM4XX_Qty").Value); //검사수량(최종)
                    oDS_PS_SD035L.SetValue("U_ColReg15", i, RecordSet01.Fields.Item("QM4XX_YN").Value); //검사여부
                    oDS_PS_SD035L.SetValue("U_ColReg16", i, RecordSet01.Fields.Item("SD040_Date").Value); //출하일자(최종)
                    oDS_PS_SD035L.SetValue("U_ColQty04", i, RecordSet01.Fields.Item("SD040_Qty").Value); //출하수량(최종)
                    oDS_PS_SD035L.SetValue("U_ColDt01", i, RecordSet01.Fields.Item("Out_Date").Value); //출하일자(신규)

                    RecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                oForm.Update();
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
            finally
            {
                oForm.Freeze(false);

                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        private void PS_SD035_MTX02()
        {
            int i;
            string Query01;
            string Out_FrDt; //출하일자(시작)
            string Out_ToDt; //출하일자(종료)
            string CardType; //거래처구분
            string ItemType; //품목구분
            string OutCls; //출고구분
            string CardCode; //수주처
            string ItemCode; //작번
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                Out_FrDt = oForm.Items.Item("Out_FrDt").Specific.Value.ToString().Trim(); //출하일자(시작)
                Out_ToDt = oForm.Items.Item("Out_ToDt").Specific.Value.ToString().Trim(); //출하일자(종료)
                CardType = oForm.Items.Item("SCardType").Specific.Value.ToString().Trim(); //거래처구분
                ItemType = oForm.Items.Item("SItemType").Specific.Value.ToString().Trim(); //품목구분
                OutCls = oForm.Items.Item("SOutCls").Specific.Value.ToString().Trim(); //출고구분
                CardCode = oForm.Items.Item("SCardCode").Specific.Value.ToString().Trim(); //수주처
                ItemCode = oForm.Items.Item("SItemCode").Specific.Value.ToString().Trim(); //작번
                
                oForm.Freeze(true);

                Query01 = "EXEC PS_SD035_03 '";
                Query01 += Out_FrDt + "','";
                Query01 += Out_ToDt + "','";
                Query01 += CardType + "','";
                Query01 += ItemType + "','";
                Query01 += OutCls + "','";
                Query01 += CardCode + "','";
                Query01 += ItemCode + "'";

                RecordSet01.DoQuery(Query01);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (RecordSet01.RecordCount == 0)
                {
                    oMat01.Clear();
                    errMessage = "";
                    throw new Exception();
                }

                for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_SD035L.InsertRecord((i));
                    }
                    oDS_PS_SD035L.Offset = i;
                    oDS_PS_SD035L.SetValue("U_LineNum", i, Convert.ToString(i + 1)); //라인번호
                    oDS_PS_SD035L.SetValue("U_ColReg01", i, RecordSet01.Fields.Item("Check").Value); //선택
                    oDS_PS_SD035L.SetValue("U_ColReg02", i, RecordSet01.Fields.Item("ItemCode").Value); //작번(품목코드)
                    oDS_PS_SD035L.SetValue("U_ColReg03", i, RecordSet01.Fields.Item("ItemName").Value); //품명
                    oDS_PS_SD035L.SetValue("U_ColReg04", i, RecordSet01.Fields.Item("CardType").Value); //거래처구분
                    oDS_PS_SD035L.SetValue("U_ColReg05", i, RecordSet01.Fields.Item("ItemType").Value); //품목구분
                    oDS_PS_SD035L.SetValue("U_ColReg06", i, RecordSet01.Fields.Item("CardCode").Value); //수주처코드
                    oDS_PS_SD035L.SetValue("U_ColReg07", i, RecordSet01.Fields.Item("CardName").Value); //수주처
                    oDS_PS_SD035L.SetValue("U_ColReg18", i, RecordSet01.Fields.Item("ORDR_LotNo").Value); //수주주문번호
                    oDS_PS_SD035L.SetValue("U_ColReg08", i, RecordSet01.Fields.Item("ORDR_Date").Value); //수주일자
                    oDS_PS_SD035L.SetValue("U_ColQty01", i, RecordSet01.Fields.Item("ORDR_Qty").Value); //수주수량
                    oDS_PS_SD035L.SetValue("U_ColReg10", i, RecordSet01.Fields.Item("PP080_Date").Value); //생산완료일자
                    oDS_PS_SD035L.SetValue("U_ColQty02", i, RecordSet01.Fields.Item("PP080_Qty").Value); //생산완료수량
                    oDS_PS_SD035L.SetValue("U_ColReg12", i, RecordSet01.Fields.Item("PP080_YN").Value); //생산완료여부
                    oDS_PS_SD035L.SetValue("U_ColReg13", i, RecordSet01.Fields.Item("QM4XX_Date").Value); //검사일자(최종)
                    oDS_PS_SD035L.SetValue("U_ColQty03", i, RecordSet01.Fields.Item("QM4XX_Qty").Value); //검사수량(최종)
                    oDS_PS_SD035L.SetValue("U_ColReg15", i, RecordSet01.Fields.Item("QM4XX_YN").Value); //검사여부
                    oDS_PS_SD035L.SetValue("U_ColReg16", i, RecordSet01.Fields.Item("SD040_Date").Value); //출하일자(최종)
                    oDS_PS_SD035L.SetValue("U_ColQty04", i, RecordSet01.Fields.Item("SD040_Qty").Value); //출하수량(최종)
                    oDS_PS_SD035L.SetValue("U_ColDt01", i, RecordSet01.Fields.Item("Out_Date").Value.ToString().Replace(".", "")); //출하일자(신규)
                    oDS_PS_SD035L.SetValue("U_ColReg19", i, RecordSet01.Fields.Item("DocEntry").Value); //문서번호(Key)

                    RecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                oForm.Update();
            }
            catch (Exception ex)
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
            finally
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }
        }

        /// <summary>
        /// 데이터 Insert
        /// </summary>
        private void PS_SD035_SaveData()
        {
            short loopCount = 0;
            string sQry = null;
            string DocEntry = null;
            string ItemCode = null;
            string ItemName = null;
            string ItemType = null;
            string CardType = null;
            string CardCode = null;
            string CardName = null;
            string ORDR_LotNo = null;
            string ORDR_Date = null;
            double ORDR_Qty = 0;
            string PP080_Date = null;
            double PP080_Qty = 0;
            string PP080_YN = null;
            string QM4XX_Date = null;
            double QM4XX_Qty = 0;
            string QM4XX_YN = null;
            string SD040_Date = null;
            double SD040_Qty = 0;
            string Out_Date = null;
            string errMessage = string.Empty;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset RecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oMat01.FlushToDataSource();
                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    if (oDS_PS_SD035L.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
                    {
                        ItemCode = oDS_PS_SD035L.GetValue("U_ColReg02", loopCount).ToString().Trim(); //품목코드(작번)
                        ItemName = oDS_PS_SD035L.GetValue("U_ColReg03", loopCount).ToString().Trim(); //품명
                        ItemType = oDS_PS_SD035L.GetValue("U_ColReg05", loopCount).ToString().Trim(); //품목구분
                        CardType = oDS_PS_SD035L.GetValue("U_ColReg04", loopCount).ToString().Trim(); //거래처구분
                        CardCode = oDS_PS_SD035L.GetValue("U_ColReg06", loopCount).ToString().Trim(); //수주처
                        CardName = oDS_PS_SD035L.GetValue("U_ColReg07", loopCount).ToString().Trim(); //수주처명
                        ORDR_LotNo = oDS_PS_SD035L.GetValue("U_ColReg18", loopCount).ToString().Trim(); //수주주문번호
                        ORDR_Date = oDS_PS_SD035L.GetValue("U_ColReg08", loopCount).ToString().Trim(); //수주일자
                        ORDR_Qty = Convert.ToDouble(oDS_PS_SD035L.GetValue("U_ColQty01", loopCount).ToString().Trim()); //수주수량
                        PP080_Date = oDS_PS_SD035L.GetValue("U_ColReg10", loopCount).ToString().Trim(); //생산완료일자(최종)
                        PP080_Qty = Convert.ToDouble(oDS_PS_SD035L.GetValue("U_ColQty02", loopCount).ToString().Trim()); //생산완료수량(누계)
                        PP080_YN = oDS_PS_SD035L.GetValue("U_ColReg12", loopCount).ToString().Trim(); //생산완료여부
                        QM4XX_Date = oDS_PS_SD035L.GetValue("U_ColReg13", loopCount).ToString().Trim(); //검사일자(최종)
                        QM4XX_Qty = Convert.ToDouble(oDS_PS_SD035L.GetValue("U_ColQty03", loopCount).ToString().Trim()); //검사수량(누계)
                        QM4XX_YN = oDS_PS_SD035L.GetValue("U_ColReg15", loopCount).ToString().Trim(); //검사여부
                        SD040_Date = oDS_PS_SD035L.GetValue("U_ColReg16", loopCount).ToString().Trim(); //출고일자(최종)
                        SD040_Qty = Convert.ToDouble(oDS_PS_SD035L.GetValue("U_ColQty04", loopCount).ToString().Trim()); //출고수량(누계)
                        Out_Date = codeHelpClass.Left(oDS_PS_SD035L.GetValue("U_ColDt01", loopCount).ToString().Trim(), 4) + "." + 
                                   codeHelpClass.Right(codeHelpClass.Left(oDS_PS_SD035L.GetValue("U_ColDt01", loopCount).ToString().Trim(), 6), 2) + "." + 
                                   codeHelpClass.Right(oDS_PS_SD035L.GetValue("U_ColDt01", loopCount).ToString().Trim(), 2); //출고일자(신규)
                        DocEntry = oDS_PS_SD035L.GetValue("U_ColReg19", loopCount).ToString().Trim(); //관리번호

                        if (string.IsNullOrEmpty(DocEntry))
                        {
                            //DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
                            sQry = "SELECT ISNULL(MAX(CONVERT(INT,DocEntry)), 0) FROM [Z_PS_SD035_01]";
                            RecordSet01.DoQuery(sQry);

                            if (Convert.ToDouble(RecordSet01.Fields.Item(0).Value.ToString().Trim()) == 0)
                            {
                                DocEntry = "1";
                            }
                            else
                            {
                                DocEntry = Convert.ToString(Convert.ToDouble(RecordSet01.Fields.Item(0).Value.ToString().Trim()) + 1);
                            }

                            sQry = "EXEC [PS_SD035_02] ";
                            sQry += "'" + DocEntry + "',"; //관리번호
                            sQry += "'" + ItemCode + "',"; //품목코드
                            sQry += "'" + ItemName + "',"; //품명
                            sQry += "'" + ItemType + "',"; //품목구분
                            sQry += "'" + CardType + "',"; //거래처구분
                            sQry += "'" + CardCode + "',"; //수주처코드
                            sQry += "'" + CardName + "',"; //수주처명
                            sQry += "'" + ORDR_LotNo + "',"; //수주주문번호
                            sQry += "'" + ORDR_Date + "',"; //수주일자
                            sQry += "'" + ORDR_Qty + "',"; //수주수량
                            sQry += "'" + PP080_Date + "',"; //생산완료일자(최종)
                            sQry += "'" + PP080_Qty + "',"; //생산완료수량(누계)
                            sQry += "'" + PP080_YN + "',"; //생산완료여부
                            sQry += "'" + QM4XX_Date + "',"; //검사일자(최종)
                            sQry += "'" + QM4XX_Qty + "',"; //검사수량(누계)
                            sQry += "'" + QM4XX_YN + "',"; //검사여부
                            sQry += "'" + SD040_Date + "',"; //출고일자(최종)
                            sQry += "'" + SD040_Qty + "',"; //출고일자(누계)
                            sQry += "'" + Out_Date + "'"; //출고일자
                        }
                        else
                        {
                            sQry = "EXEC [PS_SD035_04]";
                            sQry += "'" + DocEntry + "',"; //관리번호
                            sQry += "'" + Out_Date + "'"; //출고일자
                        }

                        RecordSet02.DoQuery(sQry);
                    }
                }

                PSH_Globals.SBO_Application.StatusBar.SetText("등록 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
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
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }
        }

        /// <summary>
        /// 데이터 Delete
        /// </summary>
        private void PS_SD035_DeleteData()
        {
            short loopCount;
            string sQry;
            string DocEntry;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oMat01.FlushToDataSource();
                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
                {
                    if (oDS_PS_SD035L.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
                    {
                        DocEntry = oDS_PS_SD035L.GetValue("U_ColReg19", loopCount).ToString().Trim(); //관리번호

                        if (string.IsNullOrEmpty(DocEntry))
                        {
                            errMessage = "관리번호가 존재하지 않아 삭제할 수 없습니다. 출하리스트 조회를 이용하십시오.";
                            throw new Exception();
                        }
                        else
                        {
                            sQry = "EXEC [PS_SD035_05]";
                            sQry += "'" + DocEntry + "'"; //관리번호
                        }

                        RecordSet01.DoQuery(sQry);
                    }
                }

                PSH_Globals.SBO_Application.StatusBar.SetText("삭제 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
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
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }
        }


        /// <summary>
        /// 리포트 출력
        /// </summary>
        private void PS_SD035_PrintReport01()
        {
            string WinTitle;
            string ReportName;
            string Out_FrDt; //출하일자(시작)
            string Out_ToDt; //출하일자(종료)
            string CardType; //거래처구분
            string ItemType; //품목구분
            string OutCls; //출고구분
            string CardCode; //수주처
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                Out_FrDt = oForm.Items.Item("Out_FrDt").Specific.Value.ToString().Trim(); //출하일자(시작)
                Out_ToDt = oForm.Items.Item("Out_ToDt").Specific.Value.ToString().Trim(); //출하일자(종료)
                CardType = oForm.Items.Item("SCardType").Specific.Value.ToString().Trim(); //거래처구분
                ItemType = oForm.Items.Item("SItemType").Specific.Value.ToString().Trim(); //품목구분
                OutCls = oForm.Items.Item("SOutCls").Specific.Value.ToString().Trim(); //출고구분
                CardCode = oForm.Items.Item("SCardCode").Specific.Value.ToString().Trim(); //수주처
                
                WinTitle = "[PS_SD035] 출하리스트";
                ReportName = "PS_SD035_01.rpt";
                //프로시저 : PS_SD035_06       

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>
                {
                    new PSH_DataPackClass("@Out_FrDt", Out_FrDt),
                    new PSH_DataPackClass("@Out_ToDt", Out_ToDt),
                    new PSH_DataPackClass("@CardType", CardType),
                    new PSH_DataPackClass("@ItemType", ItemType),
                    new PSH_DataPackClass("@OutCls", OutCls),
                    new PSH_DataPackClass("@UserSign", CardCode)
                };

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
            }
        }




        #region Raise_ItemEvent
        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	switch (pVal.EventType) {
        //		case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //			////1
        //			Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //			////2
        //			Raise_EVENT_KEY_DOWN(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //			////5
        //			Raise_EVENT_COMBO_SELECT(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_CLICK:
        //			////6
        //			Raise_EVENT_CLICK(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //			////7
        //			Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //			////8
        //			Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //			////10
        //			Raise_EVENT_VALIDATE(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //			////11
        //			Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //			////18
        //			break;
        //		////et_FORM_ACTIVATE
        //		case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //			////19
        //			break;
        //		////et_FORM_DEACTIVATE
        //		case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //			////20
        //			Raise_EVENT_RESIZE(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //			////27
        //			Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //			////3
        //			Raise_EVENT_GOT_FOCUS(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //			////4
        //			break;
        //		////et_LOST_FOCUS
        //		case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //			////17
        //			Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //	}
        //	return;
        //	Raise_ItemEvent_Error:
        //	///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_MenuEvent
        //public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	short i = 0;

        //	////BeforeAction = True
        //	if ((pVal.BeforeAction == true)) {
        //		switch (pVal.MenuUID) {
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1293":
        //				//행삭제
        //				Raise_EVENT_ROW_DELETE(ref FormUID, ref pVal, ref BubbleEvent);
        //				break;
        //			case "1281":
        //				//찾기
        //				break;
        //			case "1282":
        //				//추가
        //				break;
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				break;

        //			case "7169":
        //				//엑셀 내보내기

        //				//엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
        //				PS_SD035_AddMatrixRow(oMat01.VisualRowCount);
        //				break;

        //		}
        //	////BeforeAction = False
        //	} else if ((pVal.BeforeAction == false)) {
        //		switch (pVal.MenuUID) {
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1293":
        //				//행삭제
        //				Raise_EVENT_ROW_DELETE(ref FormUID, ref pVal, ref BubbleEvent);
        //				break;
        //			case "1281":
        //				//찾기
        //				PS_SD035_EnableFormItem();
        //				////UDO방식
        //				break;
        //			case "1282":
        //				//추가
        //				PS_SD035_EnableFormItem();
        //				////UDO방식
        //				PS_SD035_AddMatrixRow(0, ref true);
        //				////UDO방식
        //				break;
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				PS_SD035_EnableFormItem();
        //				break;
        //			case "1287":

        //				oForm.Freeze(true);
        //				oDS_PS_SD035H.SetValue("DocEntry", 0, "");

        //				for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //					oMat01.FlushToDataSource();
        //					oDS_PS_SD035H.SetValue("DocEntry", i, "");
        //					oMat01.LoadFromDataSource();
        //				}

        //				oForm.Freeze(false);
        //				break;

        //			case "7169":
        //				//엑셀 내보내기

        //				//엑셀 내보내기 이후 처리
        //				oForm.Freeze(true);
        //				oDS_PS_SD035L.RemoveRecord(oDS_PS_SD035L.Size - 1);
        //				oMat01.LoadFromDataSource();
        //				oForm.Freeze(false);
        //				break;

        //		}
        //	}
        //	return;
        //	Raise_MenuEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_FormDataEvent
        //public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	////BeforeAction = True
        //	if ((BusinessObjectInfo.BeforeAction == true)) {
        //		switch (BusinessObjectInfo.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //				////33
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //				////34
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //				////35
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //				////36
        //				break;
        //		}
        //	////BeforeAction = False
        //	} else if ((BusinessObjectInfo.BeforeAction == false)) {
        //		switch (BusinessObjectInfo.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //				////33
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //				////34
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //				////35
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //				////36
        //				break;
        //		}
        //	}
        //	return;
        //	Raise_FormDataEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_RightClickEvent
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {
        //		//        If pVal.ItemUID = "Mat01" And pVal.Row > 0 And pVal.Row <= oMat01.RowCount Then
        //		//            Dim MenuCreationParams01 As SAPbouiCOM.MenuCreationParams
        //		//            Set MenuCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        //		//            MenuCreationParams01.Type = SAPbouiCOM.BoMenuType.mt_STRING
        //		//            MenuCreationParams01.uniqueID = "MenuUID"
        //		//            MenuCreationParams01.String = "메뉴명"
        //		//            MenuCreationParams01.Enabled = True
        //		//            Call Sbo_Application.Menus.Item("1280").SubMenus.AddEx(MenuCreationParams01)
        //		//        End If
        //	} else if (pVal.BeforeAction == false) {
        //		//        If pVal.ItemUID = "Mat01" And pVal.Row > 0 And pVal.Row <= oMat01.RowCount Then
        //		//                Call Sbo_Application.Menus.RemoveEx("MenuUID")
        //		//        End If
        //	}
        //	if (pVal.ItemUID == "Mat01") {
        //		if (pVal.Row > 0) {
        //			oLastItemUID01 = pVal.ItemUID;
        //			oLastColUID01 = pVal.ColUID;
        //			oLastColRow01 = pVal.Row;
        //		}
        //	} else {
        //		oLastItemUID01 = pVal.ItemUID;
        //		oLastColUID01 = "";
        //		oLastColRow01 = 0;
        //	}
        //	return;
        //	Raise_RightClickEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ITEM_PRESSED
        //private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemUID == "PS_SD035") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}

        //		if (pVal.ItemUID == "BtnSrch01") {
        //			PS_SD035_MTX01();
        //			//대상조회
        //		}

        //		if (pVal.ItemUID == "BtnSrch02") {
        //			PS_SD035_MTX02();
        //			//출하리스트 조회
        //		}

        //		//저장
        //		if (pVal.ItemUID == "BtnSave") {
        //			PS_SD035_SaveData();
        //		}

        //		//삭제
        //		if (pVal.ItemUID == "BtnDelete") {
        //			if (SubMain.Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1")) {
        //				PS_SD035_DeleteData();
        //				PS_SD035_MTX02();
        //			} else {

        //			}
        //		}

        //		//인쇄
        //		if (pVal.ItemUID == "BtnPrint") {
        //			PS_SD035_PrintReport01();
        //		}

        //	} else if (pVal.BeforeAction == false) {
        //		if (pVal.ItemUID == "PS_SD035") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}
        //		//        If pVal.ItemUID = "1" Then
        //		//            If oForm.Mode = fm_ADD_MODE Then
        //		//                If pVal.ActionSuccess = True Then
        //		//                    Call PS_SD035_EnableFormItem
        //		//                    Call PS_SD035_AddMatrixRow(0, True) '//UDO방식일때
        //		//                End If
        //		//            ElseIf oForm.Mode = fm_UPDATE_MODE Then
        //		//            ElseIf oForm.Mode = fm_OK_MODE Then
        //		//                If pVal.ActionSuccess = True Then
        //		//                    Call PS_SD035_EnableFormItem
        //		//                End If
        //		//            End If
        //		//        End If

        //	}
        //	return;
        //	Raise_EVENT_ITEM_PRESSED_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_KEY_DOWN
        //private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	if (pVal.BeforeAction == true) {

        //		if (pVal.ItemUID == "Mat01") {

        //		//수주처(대상조회)
        //		} else if (pVal.ItemUID == "CardCode") {

        //			//탭을 눌렀을 경우만
        //			if (pVal.CharPressed == 9) {

        //				MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", "");

        //			}

        //		//수주처(출하리스트)
        //		} else if (pVal.ItemUID == "SCardCode") {

        //			//탭을 눌렀을 경우만
        //			if (pVal.CharPressed == 9) {

        //				MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "SCardCode", "");

        //			}

        //		//작번(대상조회)
        //		} else if (pVal.ItemUID == "ItemCode") {

        //			//탭을 눌렀을 경우만
        //			if (pVal.CharPressed == 9) {

        //				MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", "");

        //			}

        //		//작번(출하리스트)
        //		} else if (pVal.ItemUID == "SItemCode") {

        //			//탭을 눌렀을 경우만
        //			if (pVal.CharPressed == 9) {

        //				MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "SItemCode", "");

        //			}

        //		}

        //	} else if (pVal.BeforeAction == false) {


        //	}

        //	return;
        //	Raise_EVENT_KEY_DOWN_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_COMBO_SELECT
        //private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm.Freeze(true);
        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {

        //		if (pVal.ItemChanged == true) {

        //		}
        //	}
        //	oForm.Freeze(false);
        //	return;
        //	Raise_EVENT_COMBO_SELECT_Error:
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_CLICK
        //private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {
        //		//        If pVal.ItemUID = "Mat01" Then
        //		//            If pVal.Row > 0 Then
        //		//                Call oMat01.SelectRow(pVal.Row, True, False)
        //		//            End If
        //		//        End If
        //		if (pVal.ItemUID == "Mat01") {
        //			if (pVal.Row > 0) {
        //				oLastItemUID01 = pVal.ItemUID;
        //				oLastColUID01 = pVal.ColUID;
        //				oLastColRow01 = pVal.Row;

        //				oMat01.SelectRow(pVal.Row, true, false);
        //			}
        //		} else {
        //			oLastItemUID01 = pVal.ItemUID;
        //			oLastColUID01 = "";
        //			oLastColRow01 = 0;
        //		}
        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_CLICK_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_DOUBLE_CLICK
        //private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_DOUBLE_CLICK_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_MATRIX_LINK_PRESSED
        //private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	object oTempClass = null;
        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemUID == "Mat01") {

        //			if (pVal.ColUID == "DocEntry") {

        //				oTempClass = new PS_QM300();

        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: oTempClass.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oTempClass.LoadForm(oMat01.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value);

        //			}

        //		}
        //	} else if (pVal.BeforeAction == false) {

        //	}

        //	//UPGRADE_NOTE: oTempClass 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oTempClass = null;
        //	return;
        //	Raise_EVENT_MATRIX_LINK_PRESSED_Error:
        //	//UPGRADE_NOTE: oTempClass 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oTempClass = null;
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_VALIDATE
        //private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	string[] PS_PH030 = null;

        //	oForm.Freeze(true);
        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemChanged == true) {

        //			if ((pVal.ItemUID == "Mat01")) {

        //				oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				oMat01.AutoResizeColumns();

        //			//수주처(대상조회)
        //			} else if ((pVal.ItemUID == "CardCode")) {

        //				//UPGRADE_WARNING: oForm.Items(CardName).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("CardName").Specific.Value = MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'");

        //			//수주처(출하리스트)
        //			} else if ((pVal.ItemUID == "SCardCode")) {

        //				//UPGRADE_WARNING: oForm.Items(SCardName).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("SCardName").Specific.Value = MDC_GetData.Get_ReData("CardName", "CardCode", "[OCRD]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'");

        //			//작번(대상조회)
        //			} else if ((pVal.ItemUID == "ItemCode")) {

        //				//UPGRADE_WARNING: oForm.Items(ItemName).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("ItemName").Specific.Value = MDC_GetData.Get_ReData("ItemName", "ItemCode", "[OITM]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'");

        //			//작번(출하리스트)
        //			} else if ((pVal.ItemUID == "SItemCode")) {

        //				//UPGRADE_WARNING: oForm.Items(SItemName).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("SItemName").Specific.Value = MDC_GetData.Get_ReData("ItemName", "ItemCode", "[OITM]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'");

        //			}

        //		}

        //	} else if (pVal.BeforeAction == false) {

        //	}

        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm.Freeze(false);
        //	return;
        //	Raise_EVENT_VALIDATE_Error:
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_MATRIX_LOAD
        //private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {
        //		PS_SD035_EnableFormItem();
        //		//        Call PS_SD035_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
        //		oMat01.AutoResizeColumns();
        //	}
        //	return;
        //	Raise_EVENT_MATRIX_LOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_RESIZE
        //private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pVal = null, ref bool BubbleEvent = false)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {
        //		PS_SD035_ResizeForm();
        //		oMat01.AutoResizeColumns();
        //	}
        //	return;
        //	Raise_EVENT_RESIZE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_CHOOSE_FROM_LIST
        //private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {
        //	}
        //	return;
        //	Raise_EVENT_CHOOSE_FROM_LIST_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_GOT_FOCUS
        //private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.ItemUID == "Mat01") {
        //		if (pVal.Row > 0) {
        //			oLastItemUID01 = pVal.ItemUID;
        //			oLastColUID01 = pVal.ColUID;
        //			oLastColRow01 = pVal.Row;
        //		}
        //	} else {
        //		oLastItemUID01 = pVal.ItemUID;
        //		oLastColUID01 = "";
        //		oLastColRow01 = 0;
        //	}
        //	return;
        //	Raise_EVENT_GOT_FOCUS_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_FORM_UNLOAD
        //private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {
        //	} else if (pVal.BeforeAction == false) {
        //		SubMain.RemoveForms(oFormUniqueID);
        //		//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oForm = null;
        //		//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oMat01 = null;
        //	}
        //	return;
        //	Raise_EVENT_FORM_UNLOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ROW_DELETE
        //private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	if ((oLastColRow01 > 0)) {
        //		if (pVal.BeforeAction == true) {
        //			////행삭제전 행삭제가능여부검사
        //		} else if (pVal.BeforeAction == false) {
        //			for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
        //			}
        //			oMat01.FlushToDataSource();
        //			oDS_PS_SD035L.RemoveRecord(oDS_PS_SD035L.Size - 1);
        //			oMat01.LoadFromDataSource();
        //			if (oMat01.RowCount == 0) {
        //				PS_SD035_AddMatrixRow(0);
        //			} else {
        //				if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_SD035L.GetValue("U_CntcCode", oMat01.RowCount - 1)))) {
        //					PS_SD035_AddMatrixRow(oMat01.RowCount);
        //				}
        //			}
        //		}
        //	}
        //	return;
        //	Raise_EVENT_ROW_DELETE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion
    }
}
