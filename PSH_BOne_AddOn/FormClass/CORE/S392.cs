using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using System.Collections.Generic;

namespace PSH_BOne_AddOn.Core
{
	/// <summary>
	/// 분개
	/// </summary>
	internal class S392 : PSH_BaseClass
	{
		private SAPbouiCOM.Matrix oMat01;
		private int oMat01Row;

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="formUID"></param>
		public override void LoadForm(string formUID)
		{
			try
			{
				oForm = PSH_Globals.SBO_Application.Forms.Item(formUID);
				oForm.Freeze(true);
				SubMain.Add_Forms(this, formUID, "S392");
				oMat01 = oForm.Items.Item("76").Specific;
				S392_CreateItems();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				oForm.Update();
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// S392_CreateItems
		/// </summary>
		private void S392_CreateItems()
		{
			SAPbouiCOM.Item newItem = null;

			SAPbouiCOM.Item RdoBtn01 = null; //담당(경리부서)
			SAPbouiCOM.Item RdoBtn02 = null; //팀장(경리부서)
			SAPbouiCOM.Item RdoBtn03 = null; //사업부장(경리부서)
			SAPbouiCOM.Item RdoBtn04 = null; //전무(경리부서)
			SAPbouiCOM.Item RdoBtn05 = null; //부사장(경리부서)

			SAPbouiCOM.Item RdoBtn11 = null; //담당(품의부서)
			SAPbouiCOM.Item RdoBtn12 = null; //팀장(품의부서)
			SAPbouiCOM.Item RdoBtn13 = null; //사업부장(품의부서)
			SAPbouiCOM.Item RdoBtn14 = null; //전무(품의부서)
			SAPbouiCOM.Item RdoBtn15 = null; //부사장(품의부서)

			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//기준 아이템(취소버튼)

				//경리부서 전결용 라디오버튼
				oForm.DataSources.UserDataSources.Add("RadioBtn01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);

				//담당(경리부서) 라디오 버튼
				RdoBtn01 = oForm.Items.Add("RdoBtn01", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
				RdoBtn01.Left = oForm.Items.Item("2").Left + 85;
				RdoBtn01.Top = oForm.Items.Item("2").Top - 8;
				RdoBtn01.Height = oForm.Items.Item("2").Height;
				RdoBtn01.Width = 48;
				RdoBtn01.Specific.Caption = "담당";

				oForm.Items.Item("RdoBtn01").Specific.ValOn = "A";
				oForm.Items.Item("RdoBtn01").Specific.ValOff = "0";
				oForm.Items.Item("RdoBtn01").Specific.DataBind.SetBound(true, "", "RadioBtn01");
				oForm.Items.Item("RdoBtn01").Specific.Selected = true;

				//팀장(경리부서) 라디오버튼
				RdoBtn02 = oForm.Items.Add("RdoBtn02", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
				RdoBtn02.Left = RdoBtn01.Left + RdoBtn01.Width - 3;
				RdoBtn02.Top = RdoBtn01.Top;
				RdoBtn02.Height = RdoBtn01.Height;
				RdoBtn02.Width = 48;
				RdoBtn02.Specific.Caption = "팀장";

				oForm.Items.Item("RdoBtn02").Specific.ValOn = "B";
				oForm.Items.Item("RdoBtn02").Specific.ValOff = "0";
				oForm.Items.Item("RdoBtn02").Specific.DataBind.SetBound(true, "", "RadioBtn01");
				oForm.Items.Item("RdoBtn02").Specific.GroupWith("RdoBtn01");

				//사업부장(경리부서) 라디오버튼
				RdoBtn03 = oForm.Items.Add("RdoBtn03", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
				RdoBtn03.Left = RdoBtn02.Left + RdoBtn02.Width - 3;
				RdoBtn03.Top = RdoBtn02.Top;
				RdoBtn03.Height = RdoBtn02.Height;
				RdoBtn03.Width = 48;
				RdoBtn03.Specific.Caption = "상무";

				oForm.Items.Item("RdoBtn03").Specific.ValOn = "C";
				oForm.Items.Item("RdoBtn03").Specific.ValOff = "0";
				oForm.Items.Item("RdoBtn03").Specific.DataBind.SetBound(true, "", "RadioBtn01");
				oForm.Items.Item("RdoBtn03").Specific.GroupWith("RdoBtn01");

				//전무(경리부서) 라디오버튼
				RdoBtn04 = oForm.Items.Add("RdoBtn04", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
				RdoBtn04.Left = RdoBtn03.Left + RdoBtn03.Width - 3;
				RdoBtn04.Top = RdoBtn03.Top;
				RdoBtn04.Height = RdoBtn03.Height;
				RdoBtn04.Width = 48;
				RdoBtn04.Specific.Caption = "전무";

				oForm.Items.Item("RdoBtn04").Specific.ValOn = "D";
				oForm.Items.Item("RdoBtn04").Specific.ValOff = "0";
				oForm.Items.Item("RdoBtn04").Specific.DataBind.SetBound(true, "", "RadioBtn01");
				oForm.Items.Item("RdoBtn04").Specific.GroupWith("RdoBtn01");

				//부사장(경리부서) 라디오버튼
				RdoBtn05 = oForm.Items.Add("RdoBtn05", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
				RdoBtn05.Left = RdoBtn04.Left + RdoBtn04.Width - 3;
				RdoBtn05.Top = RdoBtn04.Top;
				RdoBtn05.Height = RdoBtn04.Height;
				RdoBtn05.Width = 58;
				RdoBtn05.Specific.Caption = "부사장";

				oForm.Items.Item("RdoBtn05").Specific.ValOn = "E";
				oForm.Items.Item("RdoBtn05").Specific.ValOff = "0";
				oForm.Items.Item("RdoBtn05").Specific.DataBind.SetBound(true, "", "RadioBtn01");
				oForm.Items.Item("RdoBtn05").Specific.GroupWith("RdoBtn01");

				oForm.DataSources.UserDataSources.Item("RadioBtn01").Value = "0";

				//품의부서 전결용 라디오버튼
				oForm.DataSources.UserDataSources.Add("RadioBtn11", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);

				//담당(품의부서) 라디오 버튼
				RdoBtn11 = oForm.Items.Add("RdoBtn11", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
				RdoBtn11.Left = oForm.Items.Item("2").Left + 85;
				RdoBtn11.Top = oForm.Items.Item("2").Top + 11;
				RdoBtn11.Height = oForm.Items.Item("2").Height;
				RdoBtn11.Width = 48;
				RdoBtn11.Specific.Caption = "담당";

				oForm.Items.Item("RdoBtn11").Specific.ValOn = "A";
				oForm.Items.Item("RdoBtn11").Specific.ValOff = "0";
				oForm.Items.Item("RdoBtn11").Specific.DataBind.SetBound(true, "", "RadioBtn11");
				oForm.Items.Item("RdoBtn11").Specific.Selected = true;

				//팀장(품의부서) 라디오버튼
				RdoBtn12 = oForm.Items.Add("RdoBtn12", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
				RdoBtn12.Left = RdoBtn11.Left + RdoBtn11.Width - 3;
				RdoBtn12.Top = RdoBtn11.Top;
				RdoBtn12.Height = RdoBtn01.Height;
				RdoBtn12.Width = 48;
				RdoBtn12.Specific.Caption = "팀장";

				oForm.Items.Item("RdoBtn12").Specific.ValOn = "B";
				oForm.Items.Item("RdoBtn12").Specific.ValOff = "0";
				oForm.Items.Item("RdoBtn12").Specific.DataBind.SetBound(true, "", "RadioBtn11");
				oForm.Items.Item("RdoBtn12").Specific.GroupWith("RdoBtn11");

				//상무(품의부서) 라디오버튼
				RdoBtn13 = oForm.Items.Add("RdoBtn13", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
				RdoBtn13.Left = RdoBtn12.Left + RdoBtn12.Width - 3;
				RdoBtn13.Top = RdoBtn12.Top;
				RdoBtn13.Height = RdoBtn12.Height;
				RdoBtn13.Width = 48;
				RdoBtn13.Specific.Caption = "상무";

				oForm.Items.Item("RdoBtn13").Specific.ValOn = "C";
				oForm.Items.Item("RdoBtn13").Specific.ValOff = "0";
				oForm.Items.Item("RdoBtn13").Specific.DataBind.SetBound(true, "", "RadioBtn11");
				oForm.Items.Item("RdoBtn13").Specific.GroupWith("RdoBtn11");

				//전무(품의부서) 라디오버튼
				RdoBtn14 = oForm.Items.Add("RdoBtn14", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
				RdoBtn14.Left = RdoBtn13.Left + RdoBtn13.Width - 3;
				RdoBtn14.Top = RdoBtn13.Top;
				RdoBtn14.Height = RdoBtn13.Height;
				RdoBtn14.Width = 48;
				RdoBtn14.Specific.Caption = "전무";

				oForm.Items.Item("RdoBtn14").Specific.ValOn = "D";
				oForm.Items.Item("RdoBtn14").Specific.ValOff = "0";
				oForm.Items.Item("RdoBtn14").Specific.DataBind.SetBound(true, "", "RadioBtn11");
				oForm.Items.Item("RdoBtn14").Specific.GroupWith("RdoBtn11");

				//부사장(품의부서) 라디오버튼
				RdoBtn15 = oForm.Items.Add("RdoBtn15", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
				RdoBtn15.Left = RdoBtn14.Left + RdoBtn14.Width - 3;
				RdoBtn15.Top = RdoBtn14.Top;
				RdoBtn15.Height = RdoBtn14.Height;
				RdoBtn15.Width = 58;
				RdoBtn15.Specific.Caption = "부사장";

				oForm.Items.Item("RdoBtn15").Specific.ValOn = "E";
				oForm.Items.Item("RdoBtn15").Specific.ValOff = "0";
				oForm.Items.Item("RdoBtn15").Specific.DataBind.SetBound(true, "", "RadioBtn11");
				oForm.Items.Item("RdoBtn15").Specific.GroupWith("RdoBtn11");

				oForm.DataSources.UserDataSources.Item("RadioBtn11").Value = "0";

				//회계전표 버튼
				newItem = oForm.Items.Add("Btn01", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
				newItem.Left = oForm.Items.Item("2").Left + 330;
				newItem.Top = oForm.Items.Item("2").Top;
				newItem.Height = oForm.Items.Item("2").Height;
				newItem.Width = 70;
				newItem.Specific.Caption = "회계 전표";

				//감가상각비 버튼
				newItem = oForm.Items.Add("Btn02", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
				newItem.Left = oForm.Items.Item("2").Left + 620;
				newItem.Top = oForm.Items.Item("2").Top;
				newItem.Height = oForm.Items.Item("2").Height;
				newItem.Width = 70;
				newItem.Specific.Caption = "감가상각비";

				//부자재불출 버튼
				newItem = oForm.Items.Add("Btn03", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
				newItem.Left = oForm.Items.Item("2").Left + 700;
				newItem.Top = oForm.Items.Item("2").Top;
				newItem.Height = oForm.Items.Item("2").Height;
				newItem.Width = 70;
				newItem.Specific.Caption = "부자재불출";

				//사업장-ComboBox
				newItem = oForm.Items.Add("Static01", SAPbouiCOM.BoFormItemTypes.it_STATIC);
				newItem.Left = oForm.Items.Item("2006").Left + 93;
				newItem.Top = oForm.Items.Item("2006").Top;
				newItem.Height = oForm.Items.Item("2006").Height;
				newItem.Width = oForm.Items.Item("2006").Width;
				newItem.Specific.Caption = "사업장";

				newItem = oForm.Items.Add("BPLId02", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
				newItem.Left = oForm.Items.Item("2000").Left + 161;
				newItem.Top = oForm.Items.Item("2000").Top;
				newItem.Height = oForm.Items.Item("2000").Height;
				newItem.Width = oForm.Items.Item("2000").Width;
				newItem.FromPane = 2;
				newItem.ToPane = 2;
				newItem.DisplayDesc = true;
				newItem.Specific.DataBind.SetBound(true, "JDT1", "U_BPLId");

				newItem = oForm.Items.Add("BPLId01", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
				newItem.Left = oForm.Items.Item("2007").Left + 93;
				newItem.Top = oForm.Items.Item("2007").Top;
				newItem.Height = oForm.Items.Item("2007").Height;
				newItem.Width = oForm.Items.Item("2007").Width + 40;
				newItem.DisplayDesc = true;
				newItem.Specific.DataBind.SetBound(true, "OJDT", "U_BPLId");

				sQry = "select BPLId, BPLName from [OBPL] order by BPLId";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					newItem.Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				newItem.Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//사업장-ComboBox
				newItem = oForm.Items.Add("Static02", SAPbouiCOM.BoFormItemTypes.it_STATIC);
				newItem.Left = oForm.Items.Item("2001").Left + 161;
				newItem.Top = oForm.Items.Item("2001").Top;
				newItem.Height = oForm.Items.Item("2001").Height;
				newItem.Width = oForm.Items.Item("2001").Width;
				newItem.FromPane = 2;
				newItem.ToPane = 2;
				newItem.Specific.Caption = "사업장";

				//거래처
				newItem = oForm.Items.Add("Static03", SAPbouiCOM.BoFormItemTypes.it_STATIC);
				newItem.Left = oForm.Items.Item("Static01").Left + 10;
				newItem.Top = oForm.Items.Item("Static01").Top;
				newItem.Height = 15;
				newItem.Width = 90;
				oForm.Items.Item("Static03").Specific.Caption = "거래처";

				//거래처코드
				newItem = oForm.Items.Add("VatBP", SAPbouiCOM.BoFormItemTypes.it_EDIT);
				newItem.Left = oForm.Items.Item("Static01").Left + 90;
				newItem.Top = oForm.Items.Item("Static01").Top;
				newItem.Height = 15;
				newItem.Width = 80;

				//거래처명
				newItem = oForm.Items.Add("VatBPName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
				newItem.Left = oForm.Items.Item("VatBP").Left + 80;
				newItem.Top = oForm.Items.Item("VatBP").Top;
				newItem.Height = 15;
				newItem.Width = 130;
				oForm.Items.Item("VatBPName").Enabled = false;

				//사업자번호
				newItem = oForm.Items.Add("VatRegNo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
				newItem.Left = oForm.Items.Item("VatBPName").Left + 130;
				newItem.Top = oForm.Items.Item("VatBPName").Top;
				newItem.Height = 15;
				newItem.Width = 100;
				oForm.Items.Item("VatRegNo").Enabled = false;

				oForm.Items.Item("Static03").LinkTo = "VatBP";

				//법정지출증빙코드
				newItem = oForm.Items.Add("Static04", SAPbouiCOM.BoFormItemTypes.it_STATIC);
				newItem.Left = oForm.Items.Item("BPLId01").Left + 10;
				newItem.Top = oForm.Items.Item("BPLId01").Top + 1;
				newItem.Height = 15;
				newItem.Width = 90;
				oForm.Items.Item("Static04").Specific.Caption = "법정지출증빙";

				//코드
				newItem = oForm.Items.Add("BillCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
				newItem.Left = oForm.Items.Item("BPLId01").Left + 90;
				newItem.Top = oForm.Items.Item("Static04").Top;
				newItem.Height = 15;
				newItem.Width = 80;
				newItem.Enabled = false;

				//명
				newItem = oForm.Items.Add("BillName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
				newItem.Left = oForm.Items.Item("BillCode").Left + 80;
				newItem.Top = oForm.Items.Item("BillCode").Top;
				newItem.Height = 15;
				newItem.Width = 130;
				newItem.Enabled = false;

				//비고
				newItem = oForm.Items.Add("BillCMT", SAPbouiCOM.BoFormItemTypes.it_EDIT);
				newItem.Left = oForm.Items.Item("BillName").Left + 130;
				newItem.Top = oForm.Items.Item("BillName").Top;
				newItem.Height = 15;
				newItem.Width = 100;

				oForm.Items.Item("Static04").LinkTo = "BillCode";

				//적용버튼
				newItem = oForm.Items.Add("BtnApply", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
				newItem.Left = oForm.Items.Item("Static04").Left;
				newItem.Top = oForm.Items.Item("Static04").Top + 20;
				newItem.Height = 20;
				newItem.Width = 60;
				newItem.Specific.Caption = "전체적용";

				newItem = oForm.Items.Add("AddonText", SAPbouiCOM.BoFormItemTypes.it_STATIC);
				newItem.Top = oForm.Items.Item("1").Top - 12;
				newItem.Left = oForm.Items.Item("1").Left;
				newItem.Height = 12;
				newItem.Width = 70;
				newItem.FontSize = 10;
				newItem.Specific.Caption = "Addon running";
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(newItem);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(RdoBtn01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(RdoBtn02);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(RdoBtn03);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(RdoBtn04);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(RdoBtn05);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(RdoBtn11);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(RdoBtn12);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(RdoBtn13);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(RdoBtn14);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(RdoBtn15);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// S392_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void S392_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			int i;

			try
			{
				oForm.Freeze(true);
				switch (oUID)
				{
					case "BPLId02":

						for (i = 1; i <= oMat01.VisualRowCount; i++)
						{
							if (oMat01Row == i)
							{
								if (!string.IsNullOrEmpty(oMat01.Columns.Item("1").Cells.Item(i).Specific.Value.ToString().Trim()))
								{
									oMat01.Columns.Item("U_BPLId").Cells.Item(i).Specific.Select(oForm.Items.Item("BPLId02").Specific.Selected.Value.ToString().Trim());
								}
							}
						}
						break;

					case "BPLId01":
						for (i = 1; i <= oMat01.VisualRowCount; i++)
						{
							if (!string.IsNullOrEmpty(oMat01.Columns.Item("1").Cells.Item(i).Specific.Value.ToString().Trim()))
							{
								oMat01.Columns.Item("U_BPLId").Cells.Item(i).Specific.Select(oForm.Items.Item("BPLId01").Specific.Selected.Value.ToString().Trim());
							}
						}
						break;
				}
				if (oUID == "76")
				{
					switch (oCol)
					{
						case "U_BPLId":
							oForm.Items.Item("BPLId02").Specific.Select(oMat01.Columns.Item("U_BPLId").Cells.Item(oRow).Specific.Value.ToString().Trim());
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
		/// S392_Form_Resize
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void S392_Form_Resize(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Items.Item("Static01").Left = oForm.Items.Item("2006").Left + 93;
				oForm.Items.Item("BPLId01").Left = oForm.Items.Item("2007").Left + 93;
				oForm.Items.Item("Static02").Top = oForm.Items.Item("2001").Top;
				oForm.Items.Item("Static02").Left = oForm.Items.Item("2001").Left + 161;
				oForm.Items.Item("BPLId02").Top = oForm.Items.Item("2000").Top;
				oForm.Items.Item("BPLId02").Left = oForm.Items.Item("2000").Left + 161;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// S392_Print_Report01
		/// </summary>
		[STAThread]
		private void S392_Print_Report01()
		{
			string TransId;
			string WinTitle;
			string ReportName;
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				TransId = oForm.Items.Item("5").Specific.Value.ToString().Trim();

				WinTitle = "회계전표 [PS_FI010]";
				ReportName = "PS_FI010_01.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();

				//Formula List
				dataPackFormula.Add(new PSH_DataPackClass("@RadioBtn01", oForm.DataSources.UserDataSources.Item("RadioBtn01").Value.ToString().Trim()));
				dataPackFormula.Add(new PSH_DataPackClass("@RadioBtn11", oForm.DataSources.UserDataSources.Item("RadioBtn11").Value.ToString().Trim()));

				//Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@TransId", TransId));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// 회계전표 클릭시 배부규칙명 업데이트
		/// </summary>
		private void S392_UpdateOCRData()
		{
			string TransId;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				TransId = oForm.Items.Item("5").Specific.Value.ToString().Trim();

				sQry = " EXEC [PS_S392_01] '" + TransId + "'"; //문서번호

                oRecordSet.DoQuery(sQry);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// 각 모드에 따른 아이템설정
		/// </summary>
		private void S392_FormItemEnabled()
		{
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("VatRegNo").Enabled = false;
					oForm.Items.Item("BillName").Enabled = false;
					oForm.Items.Item("VatBPName").Enabled = false;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("VatRegNo").Enabled = false;
					oForm.Items.Item("BillName").Enabled = false;
					oForm.Items.Item("VatBPName").Enabled = false;

				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("VatRegNo").Enabled = false;
					oForm.Items.Item("BillName").Enabled = false;
					oForm.Items.Item("VatBPName").Enabled = false;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// 필수 사항 check
		/// </summary>
		/// <returns></returns>
		private bool S392_CheckDataValid()
		{
			bool returnValue = false;
			string errMessage = string.Empty;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId01").Specific.Value.ToString().Trim(), oForm.Items.Item("6").Specific.Value.ToString().Trim().Substring(0, 6)) == false)
				{
					errMessage = "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 전기일을 확인하고, 회계부서로 문의하세요.";
					oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
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
		/// Raise_EVENT_ITEM_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string errMessage = string.Empty;
			int i;
			string sQry;
			string BPLID;
			string DocDateFr;
			string DocDateTo;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = null;

			try
			{
				oForm.Freeze(true);
				if (pVal.Before_Action == true)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (S392_CheckDataValid() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (S392_CheckDataValid() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
					}
				}
				else if (pVal.Before_Action == false)
				{
					if (pVal.ItemUID == "Btn01")
					{
						System.Threading.Thread thread = new System.Threading.Thread(S392_Print_Report01);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
						S392_UpdateOCRData(); //배부규칙 업데이트
					}
					else if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
						{
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
							PSH_Globals.SBO_Application.ActivateMenuItem("1291");
						}
					}
					else if (pVal.ItemUID == "Btn02") //감가상각비
					{
						if (string.IsNullOrEmpty(oForm.Items.Item("6").Specific.Value.ToString().Trim()))
						{
							errMessage = "전기일자는 필수입니다. 확인하세요.";
							throw new Exception();
						}
						else if (string.IsNullOrEmpty(oForm.Items.Item("BPLId01").Specific.Value.ToString().Trim()))
						{
							errMessage = "사업장은 필수입니다. 확인하세요.";
							throw new Exception();
						}

						BPLID = oForm.Items.Item("BPLId01").Specific.Value.ToString().Trim();
						DocDateFr = oForm.Items.Item("6").Specific.Value.ToString().Trim().Substring(0, 6) + "01";
						DocDateTo = oForm.Items.Item("6").Specific.Value.ToString().Trim();

						sQry = "EXEC [S392_02] '" + BPLID + "', '" + DocDateFr + "', '" + DocDateTo + "'";
						oRecordSet.DoQuery(sQry);

						if (oRecordSet.RecordCount == 0)
						{
							errMessage = "조회 결과가 없습니다. 확인하세요.";
							throw new Exception();
						}

						for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
						{
							oMat01.Columns.Item("1").Cells.Item(i + 1).Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
							if (Convert.ToDouble(oRecordSet.Fields.Item(2).Value.ToString().Trim()) > 0)
							{
								oMat01.Columns.Item("5").Cells.Item(i + 1).Specific.Value = oRecordSet.Fields.Item(2).Value.ToString().Trim();
								oMat01.Columns.Item("10002014").Cells.Item(i + 1).Specific.Value = oRecordSet.Fields.Item(1).Value.ToString().Trim();
							}
							else
							{
								oMat01.Columns.Item("6").Cells.Item(i + 1).Specific.Value = oRecordSet.Fields.Item(3).Value.ToString().Trim();
							}

							oRecordSet.MoveNext();
						}
						oMat01.AutoResizeColumns();
					}
					else if (pVal.ItemUID == "Btn03") //부자재불출
					{

						ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
						if (string.IsNullOrEmpty(oForm.Items.Item("6").Specific.Value.ToString().Trim()))
						{
							errMessage = "전기일자는 필수입니다. 확인하세요.";
							throw new Exception();
						}
						else if (string.IsNullOrEmpty(oForm.Items.Item("BPLId01").Specific.Value.ToString().Trim()))
						{
							errMessage = "사업장은 필수입니다. 확인하세요.";
							throw new Exception();
						}

						BPLID = oForm.Items.Item("BPLId01").Specific.Value.ToString().Trim();
						DocDateFr = oForm.Items.Item("6").Specific.Value.ToString().Trim().Substring(0, 6) + "01";
						DocDateTo = oForm.Items.Item("6").Specific.Value.ToString().Trim();

						sQry = "EXEC [S392_01] '" + BPLID + "', '" + DocDateFr + "', '" + DocDateTo + "'";
						oRecordSet.DoQuery(sQry);

						if (oRecordSet.RecordCount == 0)
						{
							errMessage = "조회 결과가 없습니다. 확인하세요.";
							throw new Exception();
						}

						for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
						{
							oMat01.Columns.Item("1").Cells.Item(i + 1).Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();

							if (Convert.ToDouble(oRecordSet.Fields.Item(2).Value.ToString().Trim()) != 0)
							{
								oMat01.Columns.Item("5").Cells.Item(i + 1).Specific.Value = oRecordSet.Fields.Item(2).Value.ToString().Trim();
								oMat01.Columns.Item("9").Cells.Item(i + 1).Specific.Value = "저장품에서 대체";
								oMat01.Columns.Item("10002014").Cells.Item(i + 1).Specific.Value = oRecordSet.Fields.Item(1).Value.ToString().Trim();
								//oMat01.Columns.Item("U_BPLId").Cells.Item(i + 1).Specific.Select(oForm.Items.Item("BPLId01").Specific.Selected.Value.ToString().Trim());
							}
							else
							{
								oMat01.Columns.Item("6").Cells.Item(i + 1).Specific.Value = oRecordSet.Fields.Item(3).Value.ToString().Trim();
								oMat01.Columns.Item("9").Cells.Item(i + 1).Specific.Value = "본계정에 대체";
							}

							oRecordSet.MoveNext();
							ProgressBar01.Value += 1;
							ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
						}
						oMat01.AutoResizeColumns();
					}
					else if (pVal.ItemUID == "BtnApply") //전체적용
					{
						for (i = 1; i <= oMat01.VisualRowCount; i++)
						{
							if (!string.IsNullOrEmpty(oMat01.Columns.Item("1").Cells.Item(i).Specific.Value.ToString().Trim()))
							{
								oMat01.Columns.Item("U_VatBP").Cells.Item(i).Specific.Value = oForm.Items.Item("VatBP").Specific.Value.ToString().Trim(); //거래처
								oMat01.Columns.Item("U_VatBPName").Cells.Item(i).Specific.Value = oForm.Items.Item("VatBPName").Specific.Value.ToString().Trim(); //거래처명
								oMat01.Columns.Item("U_VatRegN").Cells.Item(i).Specific.Value = oForm.Items.Item("VatRegNo").Specific.Value.ToString().Trim(); //사업자등록번호
								oMat01.Columns.Item("U_BillCode").Cells.Item(i).Specific.Value = oForm.Items.Item("BillCode").Specific.Value.ToString().Trim(); //법정증빙코드
								oMat01.Columns.Item("U_BillName").Cells.Item(i).Specific.Value = oForm.Items.Item("BillName").Specific.Value.ToString().Trim(); //법정증빙명
								oMat01.Columns.Item("U_BillCMT").Cells.Item(i).Specific.Value = oForm.Items.Item("BillCMT").Specific.Value.ToString().Trim(); //법정증빙비고
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
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
				}
				BubbleEvent = false;
			}
			finally
			{
				if (ProgressBar01 != null)
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
				}
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
				if (pVal.Before_Action == true)
				{
					if (pVal.CharPressed == 9)
					{
						if (pVal.ItemUID == "76")
						{
							if (pVal.ColUID == "U_VatBP")
							{
								if (string.IsNullOrEmpty(oMat01.Columns.Item("U_VatBP").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									PSH_Globals.SBO_Application.ActivateMenuItem("7425");
									BubbleEvent = false;
								}
							}
						}
						else if (pVal.ItemUID == "VatBP") //거래처
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("VatBP").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						else if (pVal.ItemUID == "BillCode") //법정증빙코드
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("BillCode").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
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
						if (pVal.ItemUID == "BPLId02")
						{
							S392_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}
						if (pVal.ItemUID == "BPLId01")
						{
							S392_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}
						if (pVal.ItemUID == "76" && pVal.ColUID == "U_BPLId")
						{
							S392_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
		/// CLICK 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);
				if (pVal.Before_Action == true)
				{
					if (pVal.ItemUID == "76")
					{
						oMat01Row = pVal.Row;
					}
				}
				else if (pVal.Before_Action == false)
				{
					if (pVal.ItemUID == "76" && pVal.ColUID == "U_BPLId")
					{
						if (oMat01.VisualRowCount > 1 && !string.IsNullOrEmpty(oMat01.Columns.Item("U_BPLId").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
						{
							oForm.Items.Item("BPLId02").Specific.Select(oMat01.Columns.Item("U_BPLId").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
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
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
						if (pVal.ItemUID == "VatBP")
						{
							sQry = " SELECT  CardName,";
							sQry += "         VATRegNum";
							sQry += " FROM    OCRD";
							sQry += " WHERE   CardCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							oForm.Items.Item("VatBPName").Specific.Value = oRecordSet.Fields.Item("CardName").Value.ToString().Trim();
							oForm.Items.Item("VatRegNo").Specific.Value = oRecordSet.Fields.Item("VATRegNum").Value.ToString().Trim();
						}
						else if (pVal.ItemUID == "BillCode")
						{
							sQry = " SELECT  U_CdName";
							sQry += " FROM    [@PS_SY001L]";
							sQry += " WHERE   Code ='F005'";
							sQry += "         AND U_Minor = '" + oForm.Items.Item("BillCode").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							oForm.Items.Item("BillName").Specific.Value = oRecordSet.Fields.Item("U_CdName").Value.ToString().Trim();
						}
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
					oMat01.AutoResizeColumns();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
					S392_Form_Resize(FormUID, ref pVal, ref BubbleEvent);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// Raise_RightClickEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="eventInfo"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
		{
			try
			{
				if (eventInfo.BeforeAction == true)
				{
					if (eventInfo.ItemUID == "76")
					{
						if (eventInfo.Row > 0)
						{
							oMat01Row = eventInfo.Row;
						}
					}
				}
				else if (eventInfo.BeforeAction == false)
				{
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
						case "1281": //찾기
							break;
						case "1282": //추가
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
						case "7169": //엑셀 내보내기
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1281": //찾기
							S392_FormItemEnabled();
							oForm.DataSources.UserDataSources.Item("RadioBtn01").Value = "0";
							break;
						case "1282": //추가
							S392_FormItemEnabled();
							oForm.Items.Item("BPLId01").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
							oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							oForm.DataSources.UserDataSources.Item("RadioBtn01").Value = "0";
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
							S392_FormItemEnabled();
							oForm.DataSources.UserDataSources.Item("RadioBtn01").Value = "0";
							oMat01.AutoResizeColumns();
							break;
						case "1293": //행삭제
							break;
						case "7169": //엑셀 내보내기
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
