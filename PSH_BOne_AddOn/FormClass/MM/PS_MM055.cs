using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 미입고품 입고예정일 관리
	/// </summary>
	internal class PS_MM055 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_MM055L; //등록라인

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM055.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM055_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM055");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_MM055_CreateItems();
				PS_MM055_ComboBox_Setting();
				PS_MM055_LoadCaption();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1285", false); // 복원
				oForm.EnableMenu("1284", false); // 취소
				oForm.EnableMenu("1293", false); // 행삭제
				oForm.EnableMenu("1281", false);
				oForm.EnableMenu("1282", true);
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
		/// PS_MM055_CreateItems
		/// </summary>
		private void PS_MM055_CreateItems()
		{
			try
			{
				oDS_PS_MM055L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				//입력정보조회
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("BPLID").Specific.DataBind.SetBound(true, "", "BPLID");

				//기간(품의)시작
				oForm.DataSources.UserDataSources.Add("FrDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt").Specific.DataBind.SetBound(true, "", "FrDt");

				//기간(품의)종료
				oForm.DataSources.UserDataSources.Add("ToDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt").Specific.DataBind.SetBound(true, "", "ToDt");

				//거래처구분
				oForm.DataSources.UserDataSources.Add("CardType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("CardType").Specific.DataBind.SetBound(true, "", "CardType");

				//품목구분
				oForm.DataSources.UserDataSources.Add("ItemType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("ItemType").Specific.DataBind.SetBound(true, "", "ItemType");

				//품의구분
				oForm.DataSources.UserDataSources.Add("OrdCls", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("OrdCls").Specific.DataBind.SetBound(true, "", "OrdCls");

				//품의금액
				oForm.DataSources.UserDataSources.Add("Amount", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("Amount").Specific.DataBind.SetBound(true, "", "Amount");

				//기준일자
				oForm.DataSources.UserDataSources.Add("StdDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("StdDt").Specific.DataBind.SetBound(true, "", "StdDt");

				//구매업체코드
				oForm.DataSources.UserDataSources.Add("PhsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("PhsCode").Specific.DataBind.SetBound(true, "", "PhsCode");

				//구매업체명
				oForm.DataSources.UserDataSources.Add("PhsName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("PhsName").Specific.DataBind.SetBound(true, "", "PhsName");

				//출력정보조회
				//사업장
				oForm.DataSources.UserDataSources.Add("S_BPLID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("S_BPLID").Specific.DataBind.SetBound(true, "", "S_BPLID");

				//품의구분
				oForm.DataSources.UserDataSources.Add("S_OrdCls", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("S_OrdCls").Specific.DataBind.SetBound(true, "", "S_OrdCls");

				//기준일자
				oForm.DataSources.UserDataSources.Add("S_StdDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("S_StdDt").Specific.DataBind.SetBound(true, "", "S_StdDt");

				oForm.Items.Item("StdDt").Enabled = false; //입력정보 기준일자 수정 불가처리

				oForm.Items.Item("StdDt").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
				oForm.Items.Item("S_StdDt").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.Items.Item("Amount").Click();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// 콤보박스 set
		/// </summary>
		private void PS_MM055_ComboBox_Setting()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//입력정보조회
				//사업장
				sQry = "    SELECT      BPLId AS [Code],";
				sQry += "                BPLName AS [Name]";
				sQry += " FROM       [OBPL]";
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID").Specific, sQry, "", false, false);
				oForm.Items.Item("BPLID").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//거래처구분
				sQry = "    SELECT     U_Minor AS [Code], ";
				sQry += "               U_CdName AS [Name]";
				sQry += " FROM      [@PS_SY001L]";
				sQry += " WHERE     Code = 'C100'";
				oForm.Items.Item("CardType").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CardType").Specific, sQry, "", false, false);
				oForm.Items.Item("CardType").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//품목구분
				sQry = "    SELECT     U_Minor AS [Code], ";
				sQry += "               U_CdName AS [Name]";
				sQry += " FROM      [@PS_SY001L]";
				sQry += " WHERE     Code = 'S002'";
				oForm.Items.Item("ItemType").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType").Specific, sQry, "", false, false);
				oForm.Items.Item("ItemType").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//품의구분
				sQry = "    SELECT       Code AS [Code],";
				sQry += "                 Name AS [Name]";
				sQry += " FROM        [@PSH_ORDTYP]";
				sQry += " WHERE       Code IN ('10','20','30','40')";   //4개 품의대해서만 조회
				sQry += " ORDER BY   Code";
				oForm.Items.Item("OrdCls").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("OrdCls").Specific, sQry, "", false, false);
				oForm.Items.Item("OrdCls").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//출력정보조회
				//사업장
				sQry = "    SELECT      BPLId AS [Code],";
				sQry += "                BPLName AS [Name]";
				sQry += " FROM       [OBPL]";
				dataHelpClass.Set_ComboList(oForm.Items.Item("S_BPLID").Specific, sQry, "", false, false);
				oForm.Items.Item("S_BPLID").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//품의구분
				sQry = "    SELECT       Code AS [Code],";
				sQry += "                 Name AS [Name]";
				sQry += " FROM        [@PSH_ORDTYP]";
				sQry += " WHERE       Code IN ('10','20','30','40')";   //4개 품의대해서만 조회
				sQry += " ORDER BY   Code";
				oForm.Items.Item("S_OrdCls").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("S_OrdCls").Specific, sQry, "", false, false);
				oForm.Items.Item("S_OrdCls").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//매트릭스
				//품의구분
				sQry = "    SELECT       Code AS [Code],";
				sQry += "                 Name AS [Name]";
				sQry += " FROM        [@PSH_ORDTYP]";
				sQry += " WHERE       Code IN ('10','20','30','40')";   //4개 품의대해서만 조회
				sQry += " ORDER BY   Code";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("OrdCls"), sQry, "", "");

				//자체/통합
				sQry = "    SELECT      Code,";
				sQry += "                Name ";
				sQry += " FROM       [@PSH_RETYPE]";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("OrdType"), sQry, "", "");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
		/// </summary>
		private void PS_MM055_LoadCaption()
		{
			try
			{
				oForm.Freeze(true);
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("BtnAdd").Specific.Caption = "추가";
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
				{
					oForm.Items.Item("BtnAdd").Specific.Caption = "수정";
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
		/// PS_MM055_CheckAll
		/// </summary>
		private void PS_MM055_CheckAll()
		{
			string CheckType;
			int loopCount;

			try
			{
				oForm.Freeze(true);
				CheckType = "Y";

				oMat.FlushToDataSource();

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_MM055L.GetValue("U_ColReg01", loopCount).ToString().Trim() == "N")
					{
						CheckType = "N";
						break; // TODO: might not be correct. Was : Exit For
					}
				}

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					oDS_PS_MM055L.Offset = loopCount;
					if (CheckType == "N")
					{
						oDS_PS_MM055L.SetValue("U_ColReg01", loopCount, "Y");
					}
					else
					{
						oDS_PS_MM055L.SetValue("U_ColReg01", loopCount, "N");
					}
				}

				oMat.LoadFromDataSource();
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
		/// 데이터 조회
		/// </summary>
		/// <param name="pItemUID"></param>
		private void PS_MM055_MTX01(string pItemUID)
		{
			int i;
			string BPLId;
			string FrDt;     //기간(시작)
			string ToDt;     //기간(종료)
			string CardType; //거래처구분
			string ItemType; //품목구분
			decimal Amount;  //품의금액
			string OrdCls;   //품의구분
			string StdDt;    //기준일자
			string PhsCode;  //구매업체
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				ProgressBar01.Text = "조회시작!";

				if (pItemUID == "BtnSearch1")
				{
					BPLId = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
					FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
					ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();
					CardType = oForm.Items.Item("CardType").Specific.Value.ToString().Trim();
					ItemType = oForm.Items.Item("ItemType").Specific.Value.ToString().Trim();
					Amount = Convert.ToDecimal(oForm.Items.Item("Amount").Specific.Value.ToString().Trim());
					OrdCls = oForm.Items.Item("OrdCls").Specific.Value.ToString().Trim();
					StdDt = oForm.Items.Item("StdDt").Specific.Value.ToString().Trim();
					PhsCode = oForm.Items.Item("PhsCode").Specific.Value.ToString().Trim();

					sQry = "  EXEC [PS_MM055_01] '";
					sQry += BPLId + "','";
					sQry += FrDt + "','";
					sQry += ToDt + "','";
					sQry += CardType + "','";
					sQry += ItemType + "','";
					sQry += Amount + "','";
					sQry += OrdCls + "','";
					sQry += StdDt + "','";
					sQry += PhsCode + "'";

					oRecordSet.DoQuery(sQry);

					oMat.Clear();
					oDS_PS_MM055L.Clear();
					oMat.FlushToDataSource();
					oMat.LoadFromDataSource();

					if (oRecordSet.RecordCount == 0)
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						PS_MM055_LoadCaption();
						errMessage = "조회 결과가 없습니다. 확인하세요.";
						throw new Exception();
					}

					for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
					{
						if (i + 1 > oDS_PS_MM055L.Size)
						{
							oDS_PS_MM055L.InsertRecord(i);
						}

						oMat.AddRow();
						oDS_PS_MM055L.Offset = i;

						oDS_PS_MM055L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
						oDS_PS_MM055L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("Check").Value.ToString().Trim());   //선택
						oDS_PS_MM055L.SetValue("U_ColDt01", i, oRecordSet.Fields.Item("InDueDt").Value);      //입고예정일
						oDS_PS_MM055L.SetValue("U_ColTxt01", i, oRecordSet.Fields.Item("Comment").Value.ToString().Trim()); //비고
						oDS_PS_MM055L.SetValue("U_ColReg03", i, Convert.ToDateTime(oRecordSet.Fields.Item("OrdDueDt").Value.ToString().Trim()).ToString("yyyyMMdd")); //납기일
						oDS_PS_MM055L.SetValue("U_ColReg04", i, Convert.ToDateTime(oRecordSet.Fields.Item("OrdDocDt").Value.ToString().Trim()).ToString("yyyyMMdd")); //품의일
						oDS_PS_MM055L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("MatItemCd").Value.ToString().Trim()); //자재품목코드
						oDS_PS_MM055L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("MatItemNm").Value.ToString().Trim()); //자재품명
						oDS_PS_MM055L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("MatItemSpc").Value.ToString().Trim()); //자재규격
						oDS_PS_MM055L.SetValue("U_ColReg11", i, oRecordSet.Fields.Item("OrdType").Value.ToString().Trim());   //자체/통합
						oDS_PS_MM055L.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("OrdCls").Value.ToString().Trim());    //품의구분
						oDS_PS_MM055L.SetValue("U_ColReg09", i, oRecordSet.Fields.Item("OrdCardNm").Value.ToString().Trim()); //업체명
						oDS_PS_MM055L.SetValue("U_ColReg10", i, oRecordSet.Fields.Item("OrdUnit").Value.ToString().Trim());   //단위
						oDS_PS_MM055L.SetValue("U_ColQty01", i, oRecordSet.Fields.Item("OrdQty").Value.ToString().Trim());    //수량
						oDS_PS_MM055L.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("OrdAmt").Value.ToString().Trim());    //품의금액
						oDS_PS_MM055L.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("TOrdAmt").Value.ToString().Trim());   //누적품의금액
						oDS_PS_MM055L.SetValue("U_ColSum03", i, oRecordSet.Fields.Item("SaleAmt").Value.ToString().Trim());   //수주금액
						oDS_PS_MM055L.SetValue("U_ColReg13", i, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());  //작번
						oDS_PS_MM055L.SetValue("U_ColReg14", i, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());  //품명
						oDS_PS_MM055L.SetValue("U_ColReg15", i, oRecordSet.Fields.Item("ItemSpec").Value.ToString().Trim());  //규격
						oDS_PS_MM055L.SetValue("U_ColReg19", i, oRecordSet.Fields.Item("BPLID").Value.ToString().Trim());     //사업장코드
						oDS_PS_MM055L.SetValue("U_ColReg16", i, Convert.ToDateTime(oRecordSet.Fields.Item("StdDt").Value.ToString().Trim()).ToString("yyyyMMdd")); //기준일자
						oDS_PS_MM055L.SetValue("U_ColReg17", i, oRecordSet.Fields.Item("OrdEntry").Value.ToString().Trim());  //품의문서번호
						oDS_PS_MM055L.SetValue("U_ColReg18", i, oRecordSet.Fields.Item("OrdLineID").Value.ToString().Trim()); //품의라인번호
						oRecordSet.MoveNext();

						ProgressBar01.Value += 1;
						ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
					}

					oMat.LoadFromDataSource();
					oMat.AutoResizeColumns();
				}
				else if (pItemUID == "BtnSearch2")
				{
                    if (!string.IsNullOrEmpty(oForm.Items.Item("S_StdDt").Specific.Value.ToString().Trim()))
                    {
						errMessage = "등록일자는 필수입니다.";
						throw new Exception();
                    }
					BPLId = oForm.Items.Item("S_BPLID").Specific.Value.ToString().Trim();
					OrdCls = oForm.Items.Item("S_OrdCls").Specific.Value.ToString().Trim();
					StdDt = oForm.Items.Item("S_StdDt").Specific.Value.ToString().Trim();

					sQry = "  EXEC [PS_MM055_02] '";
					sQry += BPLId + "','";
					sQry += OrdCls + "','";
					sQry += StdDt + "'";
					oRecordSet.DoQuery(sQry);

					oMat.Clear();
					oDS_PS_MM055L.Clear();
					oMat.FlushToDataSource();
					oMat.LoadFromDataSource();

					if (oRecordSet.RecordCount == 0)
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						PS_MM055_LoadCaption();
						errMessage = "조회 결과가 없습니다. 확인하세요.";
						throw new Exception();
					}

					for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
					{
						if (i + 1 > oDS_PS_MM055L.Size)
						{
							oDS_PS_MM055L.InsertRecord(i);
						}

						oMat.AddRow();
						oDS_PS_MM055L.Offset = i;

						oDS_PS_MM055L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
						oDS_PS_MM055L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("Check").Value.ToString().Trim());       //선택
						oDS_PS_MM055L.SetValue("U_ColDt01", i, oRecordSet.Fields.Item("InDueDt").Value);      //입고예정일
						oDS_PS_MM055L.SetValue("U_ColTxt01", i, oRecordSet.Fields.Item("Comment").Value.ToString().Trim());     //비고
						oDS_PS_MM055L.SetValue("U_ColReg03", i,oRecordSet.Fields.Item("OrdDueDt").Value); //납기일
						oDS_PS_MM055L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("OrdDocDt").Value); //품의일
						oDS_PS_MM055L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("MatItemCd").Value.ToString().Trim());   //자재품목코드
						oDS_PS_MM055L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("MatItemNm").Value.ToString().Trim());   //자재품명
						oDS_PS_MM055L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("MatItemSpc").Value.ToString().Trim());  //자재규격
						oDS_PS_MM055L.SetValue("U_ColReg11", i, oRecordSet.Fields.Item("OrdType").Value.ToString().Trim());     //자체/통합
						oDS_PS_MM055L.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("OrdCls").Value.ToString().Trim());      //품의구분
						oDS_PS_MM055L.SetValue("U_ColReg09", i, oRecordSet.Fields.Item("OrdCardNm").Value.ToString().Trim());   //업체명
						oDS_PS_MM055L.SetValue("U_ColReg10", i, oRecordSet.Fields.Item("OrdUnit").Value.ToString().Trim());     //단위
						oDS_PS_MM055L.SetValue("U_ColQty01", i, oRecordSet.Fields.Item("OrdQty").Value.ToString().Trim());      //수량
						oDS_PS_MM055L.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("OrdAmt").Value.ToString().Trim());      //품의금액
						oDS_PS_MM055L.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("TOrdAmt").Value.ToString().Trim());     //누적품의금액
						oDS_PS_MM055L.SetValue("U_ColSum03", i, oRecordSet.Fields.Item("SaleAmt").Value.ToString().Trim());     //수주금액
						oDS_PS_MM055L.SetValue("U_ColReg13", i, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());    //작번
						oDS_PS_MM055L.SetValue("U_ColReg14", i, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());    //품명
						oDS_PS_MM055L.SetValue("U_ColReg15", i, oRecordSet.Fields.Item("ItemSpec").Value.ToString().Trim());  //규격
						oDS_PS_MM055L.SetValue("U_ColReg19", i, oRecordSet.Fields.Item("BPLID").Value.ToString().Trim());     //사업장코드
						oDS_PS_MM055L.SetValue("U_ColReg16", i, oRecordSet.Fields.Item("StdDt").Value); //기준일자
						oDS_PS_MM055L.SetValue("U_ColReg17", i, oRecordSet.Fields.Item("OrdEntry").Value.ToString().Trim());  //품의문서번호
						oDS_PS_MM055L.SetValue("U_ColReg18", i, oRecordSet.Fields.Item("OrdLineID").Value.ToString().Trim()); //품의라인번호
						oRecordSet.MoveNext();

						ProgressBar01.Value += 1;
						ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
					}

					oMat.LoadFromDataSource();
					oMat.AutoResizeColumns();
				}
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
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// 데이터 INSERT
		/// </summary>
		/// <returns></returns>
		private bool PS_MM055_AddData()
		{
			bool ReturnValue = false;
			int loopCount;
			string InDueDt; //입고예정일
			string Comment; //비고
			string OrdDueDt;//납기일
			string OrdDocDt;//품의일
			string MatItemCd;//자재품목코드
			string MatItemNm;	//자재품명
			string MatItemSpc;	//자재규격
			string OrdType;		//자체/통합
			string OrdCls;		//품의구분
			string OrdCardNm;	//업체명
			string OrdUnit;		//단위
			double OrdQty;		//수량
			decimal OrdAmt;		//품의금액
			decimal TOrdAmt;	//누적품의금액
			decimal SaleAmt;	//수주금액
			string ItemCode;	//작번
			string ItemName;	//품명
			string BPLId;		//사업장코드
			string ItemSpec;	//규격
			string StdDt;		//기준일자
			string OrdEntry;	//품의문서번호
			string OrdLineID;	//품의라인번호
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				ProgressBar01.Text = "저장중......";

				oMat.FlushToDataSource();
				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_MM055L.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
					{
						InDueDt = oDS_PS_MM055L.GetValue("U_ColDt01", loopCount).ToString().Trim();
						Comment = oDS_PS_MM055L.GetValue("U_ColTxt01", loopCount).ToString().Trim();                          //비고
						OrdDueDt = oDS_PS_MM055L.GetValue("U_ColReg03", loopCount).ToString().Trim();                         //납기일
						OrdDocDt = oDS_PS_MM055L.GetValue("U_ColReg04", loopCount).ToString().Trim();                         //품의일
						MatItemCd = oDS_PS_MM055L.GetValue("U_ColReg05", loopCount).ToString().Trim();                        //자재품목코드
						MatItemNm = oDS_PS_MM055L.GetValue("U_ColReg06", loopCount).ToString().Trim().Replace("'", "`");      //자재품명(자재품명 중에 '를 `로 변경)
						MatItemSpc = oDS_PS_MM055L.GetValue("U_ColReg07", loopCount).ToString().Trim().Replace("'", "`");     //자재규격(자재규격 중에 '를 `로 변경)
						OrdType = oDS_PS_MM055L.GetValue("U_ColReg11", loopCount).ToString().Trim();                          //자체/통합
						OrdCls = oDS_PS_MM055L.GetValue("U_ColReg08", loopCount).ToString().Trim();                           //품의구분
						OrdCardNm = oDS_PS_MM055L.GetValue("U_ColReg09", loopCount).ToString().Trim();                        //업체명
						OrdUnit = oDS_PS_MM055L.GetValue("U_ColReg10", loopCount).ToString().Trim();                          //단위
						OrdQty = Convert.ToDouble(oDS_PS_MM055L.GetValue("U_ColQty01", loopCount).ToString().Trim());         //수량
						OrdAmt = Convert.ToDecimal(oDS_PS_MM055L.GetValue("U_ColSum01", loopCount).ToString().Trim());        //품의금액
						TOrdAmt = Convert.ToDecimal(oDS_PS_MM055L.GetValue("U_ColSum02", loopCount).ToString().Trim());       //누적품의금액
						SaleAmt = Convert.ToDecimal(oDS_PS_MM055L.GetValue("U_ColSum03", loopCount).ToString().Trim());       //수주금액
						ItemCode = oDS_PS_MM055L.GetValue("U_ColReg13", loopCount).ToString().Trim();						  //작번
						ItemName = oDS_PS_MM055L.GetValue("U_ColReg14", loopCount).ToString().Trim().Replace("'", "`");		  //품명(품명 중에 '를 `로 변경)
						ItemSpec = oDS_PS_MM055L.GetValue("U_ColReg15", loopCount).ToString().Trim().Replace("'", "`");		  //규격(규격 중에 '를 `로 변경
						BPLId = oDS_PS_MM055L.GetValue("U_ColReg19", loopCount).ToString().Trim();						      //사업장코드
						StdDt = oDS_PS_MM055L.GetValue("U_ColReg16", loopCount).ToString().Trim();						      //기준일자
						OrdEntry = oDS_PS_MM055L.GetValue("U_ColReg17", loopCount).ToString().Trim();						  //품의문서번호
						OrdLineID = oDS_PS_MM055L.GetValue("U_ColReg18", loopCount).ToString().Trim();						  //품의라인번호

						sQry = "         EXEC [PS_MM055_03] ";
						sQry += "'" + InDueDt + "',";
						sQry += "'" + Comment + "',";
						sQry += "'" + OrdDueDt + "',";
						sQry += "'" + OrdDocDt + "',";
						sQry += "'" + MatItemCd + "',";
						sQry += "'" + MatItemNm + "',";
						sQry += "'" + MatItemSpc + "',";
						sQry += "'" + OrdType + "',";
						sQry += "'" + OrdCls + "',";
						sQry += "'" + OrdCardNm + "',";
						sQry += "'" + OrdUnit + "',";
						sQry += "'" + OrdQty + "',";
						sQry += "'" + OrdAmt + "',";
						sQry += "'" + TOrdAmt + "',";
						sQry += "'" + SaleAmt + "',";
						sQry += "'" + ItemCode + "',";
						sQry += "'" + ItemName + "',";
						sQry += "'" + ItemSpec + "',";
						sQry += "'" + BPLId + "',";
						sQry += "'" + StdDt + "',";
						sQry += "'" + OrdEntry + "',";
						sQry += "'" + OrdLineID + "'";
						oRecordSet.DoQuery(sQry);
					}
				}
				PSH_Globals.SBO_Application.StatusBar.SetText("등록 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
				ReturnValue = true;
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
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return ReturnValue;
		}

		/// <summary>
		/// 기본정보 삭제 (없슴)
		/// </summary>
		private void PS_MM055_DeleteData()
		{
			int loopCount;
			string StdYM;  //기준년월
			string StdCnt; //기준회차
			string ReqNo;  //요청번호
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (oMat.VisualRowCount == 0)
				{
					errMessage = "삭제대상이 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					StdYM = oDS_PS_MM055L.GetValue("U_ColReg02", loopCount).ToString().Trim();
					StdCnt = oDS_PS_MM055L.GetValue("U_ColReg03", loopCount).ToString().Trim();
					ReqNo = oDS_PS_MM055L.GetValue("U_ColReg04", loopCount).ToString().Trim();

					sQry = "         EXEC [PS_MM055_04] ";
					sQry += "'" + StdYM + "',";
					sQry += "'" + StdCnt + "',";
					sQry += "'" + ReqNo + "'";
					oRecordSet.DoQuery(sQry);
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
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_MM055_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_MM055_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			int loopCount;
			double totalAmt = 0;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				switch (oUID)
				{
					case "Mat01":
						if (oCol == "Amount")
						{
							oMat.FlushToDataSource();

							for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
							{
								totalAmt += Convert.ToDouble(oDS_PS_MM055L.GetValue("U_ColSum03", loopCount).ToString().Trim());
							}

							oForm.Items.Item("Total").Specific.Value = Convert.ToString(totalAmt);
							oMat.LoadFromDataSource();
						}
						oMat.AutoResizeColumns();
						break;

					case "PhsCode":
						oForm.Items.Item("PhsName").Specific.Value = dataHelpClass.Get_ReData("CardName", "CardCode", "[OCRD]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", "");
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM055_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_MM055_Print_Report01()
		{
			string WinTitle;
			string ReportName;
			string BPLId;   //사업장
			string OrdCls;  //품의구분
			string StdDt;   //기준일자
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLId = oForm.Items.Item("S_BPLID").Specific.Value.ToString().Trim();
				OrdCls = oForm.Items.Item("S_OrdCls").Specific.Value.ToString().Trim();
				StdDt = oForm.Items.Item("S_StdDt").Specific.Value.ToString().Trim();

				WinTitle = "[PS_MM055] 미입고품 입고예정일 조회";
				ReportName = "PS_MM055_01.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				//Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLId));
				dataPackParameter.Add(new PSH_DataPackClass("@StdDt", DateTime.ParseExact(StdDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@OrdCls", OrdCls));

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
                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //	Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //	break;
                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
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
                //	Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //	break;
                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                //    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                //    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;
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
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
					Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
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
					if (pVal.ItemUID == "BtnAdd")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_MM055_AddData() == false)
							{
								BubbleEvent = false;
								return;
							}

							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_MM055_LoadCaption();
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_MM055_AddData() == false)
							{
								BubbleEvent = false;
								return;
							}

							PS_MM055_MTX01("BtnSearch2");
							PS_MM055_LoadCaption();
						}
					}
					else if (pVal.ItemUID == "BtnSearch1")
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						PS_MM055_LoadCaption();
						PS_MM055_MTX01(pVal.ItemUID);
					}
					else if (pVal.ItemUID == "BtnSearch2")
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
						PS_MM055_LoadCaption();
						PS_MM055_MTX01(pVal.ItemUID);
					}
					else if (pVal.ItemUID == "BtnDelete") //없슴
					{
						if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", 1, "예", "아니오") == 1)
						{
							PS_MM055_DeleteData();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_MM055_LoadCaption();
						}
					}
					else if (pVal.ItemUID == "BtnChk")
					{
						PS_MM055_CheckAll();
					}
					else if (pVal.ItemUID == "BtnPrint")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_MM055_Print_Report01);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
					}
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// Raise_EVENT_KEY_DOWN
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "PhsCode", "");
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_MM055_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// Raise_EVENT_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row > 0)
						{
							oMat.SelectRow(pVal.Row, true, false);
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
		/// Raise_EVENT_DOUBLE_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row == 0)
						{
							oMat.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
							oMat.FlushToDataSource();
						}
					}
				}
				else if (pVal.BeforeAction == false)
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
			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						PS_MM055_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
					}
				}
				else if (pVal.BeforeAction == false)
				{
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
		/// Raise_EVENT_FORM_RESIZE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					oMat.AutoResizeColumns();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM055L);
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
						case "1281": //찾기
							break;
						case "1282": //추가
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							BubbleEvent = false;
							PS_MM055_LoadCaption();
							break;
						case "1283": //삭제
							break;
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "1293": //행삭제
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1287": // 복제
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "1293": //행삭제
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
	}
}
