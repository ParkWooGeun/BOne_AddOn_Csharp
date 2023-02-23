using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using SAP.Middleware.Connector;
using PSH_BOne_AddOn.Code;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 자재 순환품 관리
	/// </summary>
	internal class PS_MM012 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;

		private SAPbouiCOM.DBDataSource oDS_PS_MM012H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_MM012L; //등록라인

		private string oCode;
		private string oYear;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM012.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM012_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM012");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
				oForm.DataBrowser.BrowseBy = "Code";

				oForm.Freeze(true);
				PS_MM012_CreateItems();
				PS_MM012_ComboBox_Setting();
				PS_MM012_FormItemEnabled();

				oForm.EnableMenu("1281", true);  // 찾기
				oForm.EnableMenu("1283", true);  // 삭제
				oForm.EnableMenu("1287", true);  // 복제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1284", false); // 취소
				oForm.EnableMenu("1293", true);  // 행삭제
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
		/// PS_MM012_CreateItems
		/// </summary>
		private void PS_MM012_CreateItems()
		{
			try
			{
				oDS_PS_MM012H = oForm.DataSources.DBDataSources.Item("@PS_MM012H");
				oDS_PS_MM012L = oForm.DataSources.DBDataSources.Item("@PS_MM012L");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;

				oDS_PS_MM012H.SetValue("U_Year", 0, DateTime.Now.ToString("yyyy"));
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM012_ComboBox_Setting
		/// </summary>
		private void PS_MM012_ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
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
		/// PS_MM012_FormItemEnabled
		/// </summary>
		private void PS_MM012_FormItemEnabled()
		{
			string YY;
			string MM;
			string DD;
			string Year_Renamed;

			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("Year").Enabled = true;
					oForm.Items.Item("Btn01").Enabled = false;
					oForm.Items.Item("Mat01").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("Year").Enabled = true;
					oForm.Items.Item("Btn01").Enabled = false;
					oForm.Items.Item("Mat01").Enabled = false;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("Year").Enabled = true;
					oForm.Items.Item("Btn01").Enabled = true;
					oForm.Items.Item("Mat01").Enabled = true;
				}

				YY = DateTime.Now.ToString("yyyy");
				MM = DateTime.Now.ToString("MM");
				DD = DateTime.Now.ToString("dd");

				Year_Renamed = oDS_PS_MM012H.GetValue("U_Year", 0).ToString().Trim();
				if (string.IsNullOrEmpty(Year_Renamed))
				{
					Year_Renamed = "0";
				}

				if (Convert.ToDouble(YY) < Convert.ToDouble(Year_Renamed))
				{
					oMat.Columns.Item("Mm01").Editable = true;
				}
				else
				{
					oMat.Columns.Item("Mm01").Editable = false;
				}

				if (YY == Year_Renamed && MM == "01")
				{
					oMat.Columns.Item("Mm01").Editable = true;
					oMat.Columns.Item("Mm02").Editable = true;
				}
				else
				{
					oMat.Columns.Item("Mm01").Editable = false;
					oMat.Columns.Item("Mm02").Editable = false;
				}
				if (YY == Year_Renamed && MM == "02")
				{
					oMat.Columns.Item("Mm02").Editable = true;
					oMat.Columns.Item("Mm03").Editable = true;
				}
				else
				{
					oMat.Columns.Item("Mm03").Editable = false;
				}
				if (YY == Year_Renamed && MM == "03")
				{
					oMat.Columns.Item("Mm03").Editable = true;
					oMat.Columns.Item("Mm04").Editable = true;
				}
				else
				{
					oMat.Columns.Item("Mm04").Editable = false;
				}
				if (YY == Year_Renamed && MM == "04")
				{
					oMat.Columns.Item("Mm04").Editable = true;
					oMat.Columns.Item("Mm05").Editable = true;
				}
				else
				{
					oMat.Columns.Item("Mm05").Editable = false;
				}
				if (YY == Year_Renamed && MM == "05")
				{
					oMat.Columns.Item("Mm05").Editable = true;
					oMat.Columns.Item("Mm06").Editable = true;
				}
				else
				{
					oMat.Columns.Item("Mm06").Editable = false;
				}
				if (YY == Year_Renamed && MM == "06")
				{
					oMat.Columns.Item("Mm06").Editable = true;
					oMat.Columns.Item("Mm07").Editable = true;
				}
				else
				{
					oMat.Columns.Item("Mm07").Editable = false;
				}
				if (YY == Year_Renamed && MM == "07")
				{
					oMat.Columns.Item("Mm07").Editable = true;
					oMat.Columns.Item("Mm08").Editable = true;
				}
				else
				{
					oMat.Columns.Item("Mm08").Editable = false;
				}
				if (YY == Year_Renamed && MM == "08")
				{
					oMat.Columns.Item("Mm08").Editable = true;
					oMat.Columns.Item("Mm09").Editable = true;
				}
				else
				{
					oMat.Columns.Item("Mm09").Editable = false;
				}
				if (YY == Year_Renamed && MM == "09")
				{
					oMat.Columns.Item("Mm09").Editable = true;
					oMat.Columns.Item("Mm10").Editable = true;
				}
				else
				{
					oMat.Columns.Item("Mm10").Editable = false;
				}
				if (YY == Year_Renamed && MM == "10")
				{
					oMat.Columns.Item("Mm10").Editable = true;
					oMat.Columns.Item("Mm11").Editable = true;
				}
				else
				{
					oMat.Columns.Item("Mm11").Editable = false;
				}
				if (YY == Year_Renamed && MM == "11")
				{
					oMat.Columns.Item("Mm11").Editable = true;
					oMat.Columns.Item("Mm12").Editable = true;
				}
				else
				{
					oMat.Columns.Item("Mm12").Editable = false;
				}
				if (YY == Year_Renamed && MM == "12")
				{
					oMat.Columns.Item("Mm12").Editable = true;
				}

				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM012_Add_MatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_MM012_Add_MatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_MM012L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_MM012L.Offset = oRow;
				oDS_PS_MM012L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM012_Copy_MatrixRow
		/// </summary>
		private void PS_MM012_Copy_MatrixRow()
		{
			int i;

			try
			{
				oDS_PS_MM012H.SetValue("Code", 0, "");
				oDS_PS_MM012H.SetValue("Name", 0, "");
				oDS_PS_MM012H.SetValue("U_Year", 0, "");

				for (i = 0; i <= oMat.VisualRowCount - 1; i++)
				{
					oMat.FlushToDataSource();
					oDS_PS_MM012L.SetValue("Code", i, "");
					oMat.LoadFromDataSource();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM012_HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_MM012_HeaderSpaceLineDel()
		{
			bool ReturnValue = false;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (string.IsNullOrEmpty(oDS_PS_MM012H.GetValue("U_Year", 0).ToString().Trim()))
				{
					errMessage = "년도는 필수입력사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_MM012H.GetValue("U_BPLId", 0).ToString().Trim()))
				{
					errMessage = "사업장은 필수입력사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					sQry = "Select Count(*) From [@PS_MM012H] Where U_BPLId = '" + oDS_PS_MM012H.GetValue("U_BPLId", 0).ToString().Trim() + "' And U_Year = '" + oDS_PS_MM012H.GetValue("U_Year", 0).ToString().Trim() + "'";
					oRecordSet.DoQuery(sQry);
					if (Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim()) > 0)
					{
						errMessage = "데이터가 있습니다. 해당 사업장과 해당 년도로 데이터를 검색하세요.";
						throw new Exception();
					}
				}

				ReturnValue = true;
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
			return ReturnValue;
		}

		/// <summary>
		/// PS_MM012_MatrixSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_MM012_MatrixSpaceLineDel()
		{
			bool ReturnValue = false;
			int i;
			string errMessage = string.Empty;

			try
			{
				oMat.FlushToDataSource();

				if (oMat.VisualRowCount == 0)
				{
					errMessage = "라인 데이터가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (i = 0; i <= oMat.VisualRowCount - 2; i++)
				{
					if (string.IsNullOrEmpty(oDS_PS_MM012L.GetValue("U_PQDocNum", i).ToString().Trim()))
					{
						errMessage = "" + i + 1 + "번 라인에 견적번호가 없습니다..행을 삭제해주세요.";
						throw new Exception();
					}
				}

				oMat.LoadFromDataSource();
				ReturnValue = true;
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
			return ReturnValue;
		}

		/// <summary>
		/// PS_MM012_Delete_EmptyRow
		/// </summary>
		private void PS_MM012_Delete_EmptyRow()
		{
			int i;

			try
			{
				oMat.FlushToDataSource();

				for (i = 0; i <= oMat.VisualRowCount - 1; i++)
				{
					if (string.IsNullOrEmpty(oDS_PS_MM012L.GetValue("U_PQDocNum", i).ToString().Trim()))
					{
						oDS_PS_MM012L.RemoveRecord(i); // Mat01에 마지막라인(빈라인) 삭제
					}
				}

				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// 본사 데이터 전송
		/// </summary>
		private bool PS_MM012_InterfaceB1toR3()
		{
			bool returnValue = false;
			string sQry;
			string Client; //클라이언트
			string ServerIP; //서버IP
			string BANFN;
			string BANPO;
			string LFDAT;
			string MEINS;
			string MENGE;
			string ZMM01;
			string ZMM02;
			string ZMM03;
			string ZMM04;
			string ZMM05;
			string ZMM06;
			string ZMM07;
			string ZMM08;
			string ZMM09;
			string ZMM10;
			string ZMM11;
			string ZMM12;
			string ZSUM;
			string errCode = string.Empty;
			string errMessage = string.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
			RfcDestination rfcDest = null;
			RfcRepository rfcRep = null;

			try
			{
				oMat.FlushToDataSource();
				Client = dataHelpClass.GetR3ServerInfo()[0];
				ServerIP = dataHelpClass.GetR3ServerInfo()[1];

				//0. 연결
				if (dataHelpClass.SAPConnection(ref rfcDest, ref rfcRep, "PSC", ServerIP, Client, "ifuser", "pdauser") == false)
				{
					errCode = "1";
					throw new Exception();
				}

				sQry = "Select * From [@PS_MM012L] Where Isnull(U_PQDocNum,'') <> '' And Code = '" + oDS_PS_MM012H.GetValue("Code", 0).ToString().Trim() + "'";
				oRecordSet.DoQuery(sQry);

				//1. SAP R3 함수 호출(매개변수 전달)
				IRfcFunction oFunction = rfcRep.CreateFunction("ZMM_INTF_GROUP2");

				while (!oRecordSet.EoF)
				{
					BANFN = oRecordSet.Fields.Item("U_E_BANFN").Value.ToString().Trim();
					BANPO = oRecordSet.Fields.Item("U_E_BNFPO").Value.ToString().Trim();
					LFDAT = oRecordSet.Fields.Item("U_DueDate").Value.ToString("yyyyMMdd").Trim();
					MEINS = oRecordSet.Fields.Item("U_Unit").Value.ToString().Trim();
					MENGE = oRecordSet.Fields.Item("U_Qty").Value.ToString().Trim();
					ZMM01 = oRecordSet.Fields.Item("U_Mm01").Value.ToString().Trim();
					ZMM02 = oRecordSet.Fields.Item("U_Mm02").Value.ToString().Trim();
					ZMM03 = oRecordSet.Fields.Item("U_Mm03").Value.ToString().Trim();
					ZMM04 = oRecordSet.Fields.Item("U_Mm04").Value.ToString().Trim();
					ZMM05 = oRecordSet.Fields.Item("U_Mm05").Value.ToString().Trim();
					ZMM06 = oRecordSet.Fields.Item("U_Mm06").Value.ToString().Trim();
					ZMM07 = oRecordSet.Fields.Item("U_Mm07").Value.ToString().Trim();
					ZMM08 = oRecordSet.Fields.Item("U_Mm08").Value.ToString().Trim();
					ZMM09 = oRecordSet.Fields.Item("U_Mm09").Value.ToString().Trim();
					ZMM10 = oRecordSet.Fields.Item("U_Mm10").Value.ToString().Trim();
					ZMM11 = oRecordSet.Fields.Item("U_Mm11").Value.ToString().Trim();
					ZMM12 = oRecordSet.Fields.Item("U_Mm12").Value.ToString().Trim();
					ZSUM = oRecordSet.Fields.Item("U_MmTot").Value.ToString().Trim();

					oFunction.SetValue("I_BANFN", BANFN);
					oFunction.SetValue("I_BNFPO", BANPO);
					oFunction.SetValue("I_LFDAT", LFDAT);
					oFunction.SetValue("I_MEINS", MEINS);
					oFunction.SetValue("I_MENGE", MENGE);
					oFunction.SetValue("I_ZMM01", ZMM01);
					oFunction.SetValue("I_ZMM02", ZMM02);
					oFunction.SetValue("I_ZMM03", ZMM03);
					oFunction.SetValue("I_ZMM04", ZMM04);
					oFunction.SetValue("I_ZMM05", ZMM05);
					oFunction.SetValue("I_ZMM06", ZMM06);
					oFunction.SetValue("I_ZMM07", ZMM07);
					oFunction.SetValue("I_ZMM08", ZMM08);
					oFunction.SetValue("I_ZMM09", ZMM09);
					oFunction.SetValue("I_ZMM10", ZMM10);
					oFunction.SetValue("I_ZMM11", ZMM11);
					oFunction.SetValue("I_ZMM12", ZMM12);
					oFunction.SetValue("I_ZSUM", ZSUM);
					oRecordSet.MoveNext();

					errCode = "2"; //SAP Function 실행 오류가 발생했을 때 에러코드로 처리하기 위해 이 위치에서 "2"를 대입
					oFunction.Invoke(rfcDest); //Function 실행

					if (oFunction.GetValue("E_MESSAGE").ToString().Trim() != "" && codeHelpClass.Left(oFunction.GetValue("E_MESSAGE").ToString().Trim(), 1) == "E") //리턴 메시지가 "S(성공)"이 아니면
					{
						errCode = "3";
						errMessage = oFunction.GetValue("E_MESSAGE").ToString();
						throw new Exception();
					}
				}

				sQry = "UPDATE [@PS_MM012H] SET U_tradate = convert(varchar(20), getdate(),120) Where U_BPLId = '" + oDS_PS_MM012H.GetValue("U_BPLId", 0).ToString().Trim() + "' And U_Year = '" + oDS_PS_MM012H.GetValue("U_Year", 0).ToString().Trim() + "'";
				oRecordSet.DoQuery(sQry);
				returnValue = true;
			}
			catch (Exception ex)
			{
				if (errCode == "1")
				{
					PSH_Globals.SBO_Application.MessageBox("풍산 SAP R3에 로그온 할 수 없습니다. 관리자에게 문의 하세요.");
				}
				else if (errCode == "2")
				{
					PSH_Globals.SBO_Application.MessageBox("RFC Function 호출 오류");
				}
				else if (errCode == "3")
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

			return returnValue;
		}

		/// <summary>
		/// PS_MM012_Qty_Check
		/// </summary>
		/// <returns></returns>
		private bool PS_MM012_Qty_Check()
		{
			bool ReturnValue = false;
			int i;
			string ItemCode;
			double Qty;
			double MmTot;
			string errMessage = string.Empty;

			try
			{
				for (i = 1; i <= oMat.VisualRowCount - 1; i++)
				{
					Qty = Convert.ToDouble(oMat.Columns.Item("Qty").Cells.Item(i).Specific.Value.ToString().Trim());
					MmTot = Convert.ToDouble(oMat.Columns.Item("MmTot").Cells.Item(i).Specific.Value.ToString().Trim());

					if (Qty < MmTot)
					{
						ItemCode = oMat.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim();
						errMessage = "(" + ItemCode + ") 구매요청 수량보다 초과할 수 없습니다.확인바랍니다.";
						throw new Exception();
					}
				}

				ReturnValue = true;
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
			return ReturnValue;
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
				//case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
				//	Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				case SAPbouiCOM.BoEventTypes.et_CLICK: //6
					Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
				//	Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//	break;
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
		/// ITEM_PRESSED 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string Code;

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_MM012_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_MM012_MatrixSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_MM012_Qty_Check() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
							{
								Code = oDS_PS_MM012H.GetValue("U_BPLId", 0).ToString().Trim() + oDS_PS_MM012H.GetValue("U_Year", 0).ToString().Trim();
								oDS_PS_MM012H.SetValue("Code", 0, Code);
								oDS_PS_MM012H.SetValue("Name", 0, Code);
							}

							PS_MM012_Delete_EmptyRow();
							oCode = oDS_PS_MM012H.GetValue("Code", 0).ToString().Trim();
						}

						oYear = oDS_PS_MM012H.GetValue("U_Year", 0).ToString().Trim();
					}
					else if (pVal.ItemUID == "Btn01")
					{
						PS_MM012_InterfaceB1toR3();
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
						{
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
							PSH_Globals.SBO_Application.ActivateMenuItem("1281");
							oDS_PS_MM012H.SetValue("Code", 0, oCode);
							oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}

						PS_MM012_FormItemEnabled();
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
			try
			{
				if (pVal.Before_Action == true)
				{
					if (pVal.CharPressed == 9)
					{
						if (pVal.ItemUID == "Mat01" && pVal.ColUID == "PQDocNum")
						{
							if (string.IsNullOrEmpty(oMat.Columns.Item("PQDocNum").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
							{
								PS_MM013 ChildForm01 = new PS_MM013();
								ChildForm01.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
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
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row > 0)
						{
							oMat.SelectRow(pVal.Row, true, false);
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
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			double MmTot;
			double Mm01;
			double Mm02;
			double Mm03;
			double Mm04;
			double Mm05;
			double Mm06;
			double Mm07;
			double Mm08;
			double Mm09;
			double Mm10;
			double Mm11;
			double Mm12;

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemChanged == true)
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							oForm.Items.Item("Btn01").Enabled = false;
						}

						if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "PQDocNum")
							{
								if ((pVal.Row == oMat.RowCount || oMat.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat.Columns.Item("PQDocNum").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									oMat.FlushToDataSource();
									oDS_PS_MM012L.InsertRecord(pVal.Row);
									oDS_PS_MM012L.SetValue("U_LineNum", pVal.Row, Convert.ToString(pVal.Row + 1));
									oMat.LoadFromDataSource();
								}
							}
							else if (pVal.ColUID == "Mm01" || pVal.ColUID == "Mm02" || pVal.ColUID == "Mm03" || pVal.ColUID == "Mm04" || pVal.ColUID == "Mm05"
									 || pVal.ColUID == "Mm06" || pVal.ColUID == "Mm07" || pVal.ColUID == "Mm08" || pVal.ColUID == "Mm09" || pVal.ColUID == "Mm10"
									 || pVal.ColUID == "Mm11" || pVal.ColUID == "Mm12")
							{
								if (string.IsNullOrEmpty(oMat.Columns.Item("Mm01").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									Mm01 = 0;
								}
								else
								{
									Mm01 = Convert.ToDouble(oMat.Columns.Item("Mm01").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								}
								if (string.IsNullOrEmpty(oMat.Columns.Item("Mm02").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									Mm02 = 0;
								}
								else
								{
									Mm02 = Convert.ToDouble(oMat.Columns.Item("Mm02").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								}
								if (string.IsNullOrEmpty(oMat.Columns.Item("Mm03").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									Mm03 = 0;
								}
								else
								{
									Mm03 = Convert.ToDouble(oMat.Columns.Item("Mm03").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								}
								if (string.IsNullOrEmpty(oMat.Columns.Item("Mm04").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									Mm04 = 0;
								}
								else
								{
									Mm04 = Convert.ToDouble(oMat.Columns.Item("Mm04").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								}
								if (string.IsNullOrEmpty(oMat.Columns.Item("Mm05").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									Mm05 = 0;
								}
								else
								{
									Mm05 = Convert.ToDouble(oMat.Columns.Item("Mm05").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								}
								if (string.IsNullOrEmpty(oMat.Columns.Item("Mm06").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									Mm06 = 0;
								}
								else
								{
									Mm06 = Convert.ToDouble(oMat.Columns.Item("Mm06").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								}
								if (string.IsNullOrEmpty(oMat.Columns.Item("Mm07").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									Mm07 = 0;
								}
								else
								{
									Mm07 = Convert.ToDouble(oMat.Columns.Item("Mm07").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								}
								if (string.IsNullOrEmpty(oMat.Columns.Item("Mm08").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									Mm08 = 0;
								}
								else
								{
									Mm08 = Convert.ToDouble(oMat.Columns.Item("Mm08").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								}
								if (string.IsNullOrEmpty(oMat.Columns.Item("Mm09").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									Mm09 = 0;
								}
								else
								{
									Mm09 = Convert.ToDouble(oMat.Columns.Item("Mm09").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								}
								if (string.IsNullOrEmpty(oMat.Columns.Item("Mm10").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									Mm10 = 0;
								}
								else
								{
									Mm10 = Convert.ToDouble(oMat.Columns.Item("Mm10").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								}
								if (string.IsNullOrEmpty(oMat.Columns.Item("Mm11").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									Mm11 = 0;
								}
								else
								{
									Mm11 = Convert.ToDouble(oMat.Columns.Item("Mm11").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								}
								if (string.IsNullOrEmpty(oMat.Columns.Item("Mm12").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									Mm12 = 0;
								}
								else
								{
									Mm12 = Convert.ToDouble(oMat.Columns.Item("Mm12").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								}

								MmTot = Mm01 + Mm02 + Mm03 + Mm04 + Mm05 + Mm06 + Mm07 + Mm08 + Mm09 + Mm10 + Mm11 + Mm12;
								oMat.Columns.Item("MmTot").Cells.Item(pVal.Row).Specific.Value = Convert.ToString(MmTot);
							}

							oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
		/// Raise_EVENT_MATRIX_LOAD
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_MM012_Add_MatrixRow(oMat.RowCount, false);
					PS_MM012_FormItemEnabled();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM012H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM012L);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// Raise_EVENT_ROW_DELETE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			int i;

			try
			{
				if (oMat.RowCount != oMat.VisualRowCount)
				{
					for (i = 0; i <= oMat.VisualRowCount - 1; i++)
					{
						oMat.Columns.Item("LineNum").Cells.Item(i + 1).Specific.Value = i + 1;
					}

					oMat.FlushToDataSource();
					oDS_PS_MM012L.RemoveRecord(oDS_PS_MM012L.Size - 1);
					oMat.Clear();
					oMat.LoadFromDataSource();
					if (!string.IsNullOrEmpty(oMat.Columns.Item("PQDocNum").Cells.Item(oMat.RowCount).Specific.Value.ToString().Trim()))
					{
						PS_MM012_Add_MatrixRow(oMat.RowCount, false);
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
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1283": //삭제
							if (PSH_Globals.SBO_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2)
							{
								BubbleEvent = false;
								return;
							}
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
							PS_MM012_FormItemEnabled();
							break;
						case "1282": //추가
							PS_MM012_FormItemEnabled();
							PS_MM012_Add_MatrixRow(0, true);
							break;
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1287": // 복제
							PS_MM012_Copy_MatrixRow();
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							PS_MM012_FormItemEnabled();
							break;
						case "1293": //행삭제
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
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
