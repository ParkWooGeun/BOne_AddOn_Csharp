//using Microsoft.VisualBasic;
//using Microsoft.VisualBasic.Compatibility;
//using System;
//using System.Collections;
//using System.Data;
//using System.Diagnostics;
//using System.Drawing;
//using System.Windows.Forms;
// // ERROR: Not supported in C#: OptionDeclaration
//namespace MDC_PS_Addon
//{
//	static class MDC_PS_Common
//	{

//		public static void ConnectODBC()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			MDC_Globals.gParam_DataBase = SubMain.Sbo_Application.Company.DatabaseName;
//			MDC_Globals.gParam_Server = SubMain.Sbo_Application.Company.ServerName;
//			//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			MDC_Globals.gParam_DBID = MDC_PS_Common.GetValue("EXEC Profile_SELECT 'SERVERINFO'", 6, 1);
//			//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			MDC_Globals.gParam_DBPW = MDC_PS_Common.GetValue("EXEC Profile_SELECT 'SERVERINFO'", 7, 1);

//			if (Strings.Right(MDC_Globals.gParam_Server, 3) == "223") {
//				MDC_Globals.gParam_ODBC = "PSHERP_TEST";
//			} else {
//				MDC_Globals.gParam_ODBC = "MDCERP";
//			}

//			MDC_Globals.ZG_CRWDSN = "PROVIDER=MSDASQL;DSN=" + MDC_Globals.gParam_ODBC + ";DATABASE=" + MDC_Globals.gParam_DataBase + ";UID=" + MDC_Globals.gParam_DBID + ";PWD=" + MDC_Globals.gParam_DBPW + ";";

//			////ZG_CRWDSN = "PROVIDER=SQLOLEDB;Data Source=" & gParam_Server & ";Initial Catalog=" & gParam_DataBase & ";User ID=" & gParam_DBID & ";Password=" & gParam_DBPW & ";"
//			MDC_Globals.g_ERPDMS = new ADODB.Connection();
//			MDC_Globals.g_ERPDMS.ConnectionTimeout = 60;
//			MDC_Globals.g_ERPDMS.CommandTimeout = 120;
//			MDC_Globals.g_ERPDMS.CursorLocation = ADODB.CursorLocationEnum.adUseClient;
//			MDC_Globals.g_ERPDMS.Open(MDC_Globals.ZG_CRWDSN);
//			if (Err().Number != 0) {
//				SubMain.Sbo_Application.SetStatusBarMessage("ODBC데이터베이스 연결에 실패하였습니다. ODBC설정을 확인하십시오!! ", SAPbouiCOM.BoMessageTime.bmt_Short, false);
//			}
//        }

//		public static void Combo_ValidValues_SetValueItem(ref SAPbouiCOM.ComboBox Combo, string FormUID, string ItemUID, bool EmptyValue = false)
//		{
//			object i = null;
//			string Query01 = null;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			Query01 = "SELECT VALUE,DESCRIPTION FROM COMBO_VALIDVALUES WHERE FORMUID = '" + FormUID + "' AND ITEMUID = '" + ItemUID + "'";
//			RecordSet01.DoQuery(Query01);
//			if ((RecordSet01.RecordCount > 0)) {
//				for (i = 1; i <= Combo.ValidValues.Count; i++) {
//					Combo.ValidValues.Remove((0));
//				}
//				if (EmptyValue == true) {
//					Combo.ValidValues.Add("", "");
//				}
//				for (i = 1; i <= RecordSet01.RecordCount; i++) {
//					Combo.ValidValues.Add(RecordSet01.Fields.Item(0).Value, RecordSet01.Fields.Item(1).Value);
//					RecordSet01.MoveNext();
//				}
//			}
//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//		}

//		public static void Combo_ValidValues_SetValueColumn(ref SAPbouiCOM.Column Column, string FormUID, string ItemUID, string ColumnUID, bool EmptyValue = false)
//		{
//			object i = null;
//			string Query01 = null;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			Query01 = "SELECT VALUE,DESCRIPTION FROM COMBO_VALIDVALUES WHERE FORMUID = '" + FormUID + "' AND ITEMUID = '" + ItemUID + "' AND COLUMNUID = '" + ColumnUID + "'";
//			RecordSet01.DoQuery(Query01);
//			if ((RecordSet01.RecordCount > 0)) {
//				for (i = 1; i <= Column.ValidValues.Count; i++) {
//					Column.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
//				}
//				if (EmptyValue == true) {
//					Column.ValidValues.Add("", "");
//				}
//				for (i = 1; i <= RecordSet01.RecordCount; i++) {
//					Column.ValidValues.Add(RecordSet01.Fields.Item(0).Value, RecordSet01.Fields.Item(1).Value);
//					RecordSet01.MoveNext();
//				}
//			}
//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//		}

//		public static void DoQuery(string Query01)
//		{
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			RecordSet01.DoQuery(Query01);
//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//		}

//		public static object GetValue(string Query01, int FieldCount = 0, int RecordCount = 0)
//		{
//			object functionReturnValue = null;
//			int i = 0;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			RecordSet01.DoQuery(Query01);
//			if ((RecordSet01.RecordCount > 0)) {
//				RecordSet01.MoveFirst();
//				if ((RecordCount == 0)) {
//					RecordCount = 1;
//				}
//				for (i = 1; i <= RecordCount; i++) {
//					//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					//UPGRADE_WARNING: GetValue 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					functionReturnValue = RecordSet01.Fields.Item(FieldCount).Value;
//					RecordSet01.MoveNext();
//				}
//			} else {
//				//UPGRADE_WARNING: GetValue 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				functionReturnValue = "";
//			}
//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			return functionReturnValue;
//		}

//		public static void ActiveUserDefineValue(ref SAPbouiCOM.Form oForm01, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent, string ItemUID, string ColumnUID = "")
//		{
//			if (string.IsNullOrEmpty(ColumnUID)) {
//				if (pval.ItemUID == ItemUID) {
//					if (pval.CharPressed == Convert.ToDouble("9")) {
//						//UPGRADE_WARNING: oForm01.Items(ItemUID).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (string.IsNullOrEmpty(oForm01.Items.Item(ItemUID).Specific.VALUE)) {
//							SubMain.Sbo_Application.ActivateMenuItem(("7425"));
//							BubbleEvent = false;
//						}
//					}
//				}
//			} else {
//				if (pval.ItemUID == ItemUID) {
//					if (pval.ColUID == ColumnUID) {
//						if (pval.CharPressed == Convert.ToDouble("9")) {
//							//UPGRADE_WARNING: oForm01.Items().Specific.Columns 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (string.IsNullOrEmpty(oForm01.Items.Item(ItemUID).Specific.Columns(ColumnUID).Cells(pval.Row).Specific.VALUE)) {
//								SubMain.Sbo_Application.ActivateMenuItem(("7425"));
//								BubbleEvent = false;
//							}
//						}
//					}
//				}
//			}
//		}

//		public static void ActiveUserDefineValueAlways(ref SAPbouiCOM.Form oForm01, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent, string ItemUID, string ColumnUID = "")
//		{
//			if (string.IsNullOrEmpty(ColumnUID)) {
//				if (pval.ItemUID == ItemUID) {
//					if (pval.CharPressed == Convert.ToDouble("9")) {
//						//UPGRADE_WARNING: oForm01.Items(ItemUID).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (string.IsNullOrEmpty(oForm01.Items.Item(ItemUID).Specific.VALUE)) {
//							SubMain.Sbo_Application.ActivateMenuItem(("7425"));
//							BubbleEvent = false;
//						}
//					} else {
//						SubMain.Sbo_Application.ActivateMenuItem(("7425"));
//						BubbleEvent = false;
//					}
//				}
//			} else {
//				if (pval.ItemUID == ItemUID) {
//					if (pval.ColUID == ColumnUID) {
//						if (pval.CharPressed == Convert.ToDouble("9")) {
//							//UPGRADE_WARNING: oForm01.Items().Specific.Columns 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (string.IsNullOrEmpty(oForm01.Items.Item(ItemUID).Specific.Columns(ColumnUID).Cells(pval.Row).Specific.VALUE)) {
//								SubMain.Sbo_Application.ActivateMenuItem(("7425"));
//								BubbleEvent = false;
//							}
//						} else {
//							SubMain.Sbo_Application.ActivateMenuItem(("7425"));
//							BubbleEvent = false;
//						}
//					}
//				}
//			}
//		}

//		public static string GetItem_UnWeight(string ItemCode)
//		{
//			string functionReturnValue = null;
//			string Query01 = null;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			Query01 = "SELECT U_UnWeight FROM [OITM] WHERE ItemCode = '" + ItemCode + "'";
//			RecordSet01.DoQuery(Query01);
//			if ((RecordSet01.RecordCount == 0)) {
//				functionReturnValue = "";
//			} else {
//				//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				functionReturnValue = RecordSet01.Fields.Item(0).Value;
//			}
//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			return functionReturnValue;
//		}

//		public static string GetItem_ItmBsort(string ItemCode)
//		{
//			string functionReturnValue = null;
//			string Query01 = null;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			Query01 = "SELECT U_ItmBsort FROM [OITM] WHERE ItemCode = '" + ItemCode + "'";
//			RecordSet01.DoQuery(Query01);
//			if ((RecordSet01.RecordCount == 0)) {
//				functionReturnValue = "";
//			} else {
//				//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				functionReturnValue = RecordSet01.Fields.Item(0).Value;
//			}
//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			return functionReturnValue;
//		}

//		public static string GetItem_SbasUnit(string ItemCode)
//		{
//			string functionReturnValue = null;
//			string Query01 = null;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			Query01 = "SELECT U_SBasUnit FROM [OITM] WHERE ItemCode = '" + ItemCode + "'";
//			RecordSet01.DoQuery(Query01);
//			if ((RecordSet01.RecordCount == 0)) {
//				functionReturnValue = "";
//			} else {
//				//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				functionReturnValue = RecordSet01.Fields.Item(0).Value;
//			}
//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			return functionReturnValue;
//		}

//		public static string GetItem_ObasUnit(string ItemCode)
//		{
//			string functionReturnValue = null;
//			string Query01 = null;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			Query01 = "SELECT U_OBasUnit FROM [OITM] WHERE ItemCode = '" + ItemCode + "'";
//			RecordSet01.DoQuery(Query01);
//			if ((RecordSet01.RecordCount == 0)) {
//				functionReturnValue = "";
//			} else {
//				//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				functionReturnValue = RecordSet01.Fields.Item(0).Value;
//			}
//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			return functionReturnValue;
//		}

//		public static string GetItem_Unit1(string ItemCode)
//		{
//			string functionReturnValue = null;
//			string Query01 = null;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			Query01 = "SELECT U_UnitQ1 FROM [OITM] WHERE ItemCode = '" + ItemCode + "'";
//			RecordSet01.DoQuery(Query01);
//			if ((RecordSet01.RecordCount == 0)) {
//				functionReturnValue = "";
//			} else {
//				//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				functionReturnValue = RecordSet01.Fields.Item(0).Value;
//			}
//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			return functionReturnValue;
//		}

//		public static string GetItem_Spec1(string ItemCode)
//		{
//			string functionReturnValue = null;
//			string Query01 = null;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			Query01 = "SELECT U_Spec1 FROM [OITM] WHERE ItemCode = '" + ItemCode + "'";
//			RecordSet01.DoQuery(Query01);
//			if ((RecordSet01.RecordCount == 0)) {
//				functionReturnValue = "";
//			} else {
//				//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				functionReturnValue = RecordSet01.Fields.Item(0).Value;
//			}
//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			return functionReturnValue;
//		}

//		public static string GetItem_Spec2(string ItemCode)
//		{
//			string functionReturnValue = null;
//			string Query01 = null;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			Query01 = "SELECT U_Spec2 FROM [OITM] WHERE ItemCode = '" + ItemCode + "'";
//			RecordSet01.DoQuery(Query01);
//			if ((RecordSet01.RecordCount == 0)) {
//				functionReturnValue = "";
//			} else {
//				//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				functionReturnValue = RecordSet01.Fields.Item(0).Value;
//			}
//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			return functionReturnValue;
//		}

//		public static string GetItem_Spec3(string ItemCode)
//		{
//			string functionReturnValue = null;
//			string Query01 = null;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			Query01 = "SELECT U_Spec3 FROM [OITM] WHERE ItemCode = '" + ItemCode + "'";
//			RecordSet01.DoQuery(Query01);
//			if ((RecordSet01.RecordCount == 0)) {
//				functionReturnValue = "";
//			} else {
//				//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				functionReturnValue = RecordSet01.Fields.Item(0).Value;
//			}
//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			return functionReturnValue;
//		}

//		public static string GetItem_ManBtchNum(string ItemCode)
//		{
//			string functionReturnValue = null;
//			string Query01 = null;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			Query01 = "SELECT ManBtchNum FROM [OITM] WHERE ItemCode = '" + ItemCode + "'";
//			RecordSet01.DoQuery(Query01);
//			if ((RecordSet01.RecordCount == 0)) {
//				functionReturnValue = "";
//			} else {
//				//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				functionReturnValue = RecordSet01.Fields.Item(0).Value;
//			}
//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			return functionReturnValue;
//		}

//		public static string GetItem_TradeType(string ItemCode)
//		{
//			string functionReturnValue = null;
//			string Query01 = null;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			Query01 = "SELECT U_TradeType FROM [OITM] WHERE ItemCode = '" + ItemCode + "'";
//			RecordSet01.DoQuery(Query01);
//			if ((RecordSet01.RecordCount == 0)) {
//				functionReturnValue = "";
//			} else {
//				//UPGRADE_WARNING: RecordSet01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				functionReturnValue = RecordSet01.Fields.Item(0).Value;
//			}
//			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			RecordSet01 = null;
//			return functionReturnValue;
//		}

//		public static void SBO_SetBackOrderFunction(ref SAPbouiCOM.Form oForm01)
//		{
//			object oRecordset01 = null;
//			 // ERROR: Not supported in C#: OnErrorStatement


//			SAPbouiCOM.Matrix oMat01 = null;
//			oMat01 = oForm01.Items.Item("38").Specific;
//			if ((oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
//				return;
//			}

//			int i = 0;
//			string Query01 = null;
//			SAPbobsCOM.Recordset RecordSet01 = null;
//			string BaseType = null;
//			string BaseTable = null;
//			int BaseEntry = 0;
//			int BaseLine = 0;
//			if ((oMat01.VisualRowCount > 1)) {
//				////선행작업의 총중량 - 현재 작업에서 생성된 중량을 뺀값을 구함
//				RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//				for (i = 1; i <= oMat01.RowCount - 1; i++) {
//					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					BaseType = oMat01.Columns.Item("43").Cells.Item(i).Specific.VALUE;
//					if ((BaseType == "-1")) {
//						goto Continue_Renamed;
//					}
//					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					BaseEntry = oMat01.Columns.Item("45").Cells.Item(i).Specific.VALUE;
//					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					BaseLine = oMat01.Columns.Item("46").Cells.Item(i).Specific.VALUE;
//					////판매오더
//					if ((BaseType == "17")) {
//						BaseTable = "RDR";
//					////판매견적
//					} else if ((BaseType == "23")) {
//						BaseTable = "QUT";
//					////납품
//					} else if ((BaseType == "15")) {
//						BaseTable = "DLN";
//					////판매반품
//					} else if ((BaseType == "16")) {
//						BaseTable = "RDN";
//					////AR송장
//					} else if ((BaseType == "13")) {
//						BaseTable = "INV";
//					////AR대변메모
//					} else if ((BaseType == "14")) {
//						BaseTable = "RIN";
//					////구매오더
//					} else if ((BaseType == "22")) {
//						BaseTable = "POR";
//					////입고PO
//					} else if ((BaseType == "20")) {
//						BaseTable = "PDN";
//					////구매반품
//					} else if ((BaseType == "21")) {
//						BaseTable = "RPD";
//					////AP송장
//					} else if ((BaseType == "18")) {
//						BaseTable = "PCH";
//					////AP대변메모
//					} else if ((BaseType == "19")) {
//						BaseTable = "RPC";
//					} else {
//						SubMain.Sbo_Application.MessageBox("화면캡쳐후 관리자에게 문의바랍니다.");
//						return;
//					}
//					Query01 = " PS_SBO_GETQUANTITY '" + BaseType + "','" + BaseTable + "','" + BaseEntry + "','" + BaseLine + "'";
//					RecordSet01.DoQuery(Query01);
//					//UPGRADE_WARNING: oMat01.Columns(U_Qty).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oMat01.Columns.Item("U_Qty").Cells.Item(i).Specific.VALUE = System.Math.Round(RecordSet01.Fields.Item(0).Value, 0);
//					//UPGRADE_WARNING: oMat01.Columns(11).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					oMat01.Columns.Item("11").Cells.Item(i).Specific.VALUE = System.Math.Round(RecordSet01.Fields.Item(1).Value, 2);
//					oMat01.Columns.Item("1").Cells.Item(oMat01.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//					Continue_Renamed:
//				}
//				//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				RecordSet01 = null;
//			}

//			return;
//			SBO_SetBackOrderFunction_Error:

//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			MDC_Com.MDC_GF_Message(ref "SBO_SetBackOrderFunction_Error:" + Err().Number + " - " + Err().Description, ref "E");

//		}

//		// 아이템 네임에 작은 따옴표 추가	
//		public static string Make_ItemName(string ItemName)
//		{
//			string functionReturnValue = null;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			int i = 0;
//			string TempItemName = null;

//			TempItemName = "";
//			for (i = 1; i <= Strings.Len(ItemName); i++) {
//				TempItemName = TempItemName + Strings.Mid(ItemName, i, 1);
//				if (Strings.Mid(ItemName, i, 1) == "'") {
//					TempItemName = TempItemName + "'";
//				}
//			}

//			functionReturnValue = Strings.Trim(TempItemName);
//			return functionReturnValue;
//			Make_ItemName_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			TempItemName = "";
//			MDC_Com.MDC_GF_Message(ref "User_BPLId_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			return functionReturnValue;
//		}

//		// 아이디별 사업장 선택
//		public static string User_BPLId()
//		{
//			string functionReturnValue = null;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordset01 = null;

//			oRecordset01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "Select Branch From [OUSR] Where USER_CODE = '" + Strings.Trim(SubMain.Sbo_Company.UserName) + "'";
//			oRecordset01.DoQuery(sQry);

//			functionReturnValue = Strings.Trim(oRecordset01.Fields.Item(0).Value);
//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			return functionReturnValue;
//			User_BPLId_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			functionReturnValue = Convert.ToString(0);
//			MDC_Com.MDC_GF_Message(ref "User_BPLId_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			return functionReturnValue;
//		}

//		// 아이디별 창고 선택 [기본창고 1, 외주가공 8, 임가공 9]
//		public static string User_WhsCode(string Gbn)
//		{
//			string functionReturnValue = null;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordset01 = null;

//			oRecordset01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "Select a.WhsCode From [OWHS] a Inner Join [OUSR] b On a.BPLid = b.Branch Where b.USER_CODE = '" + Strings.Trim(SubMain.Sbo_Company.UserName) + "' ";
//			sQry = sQry + "And LEFT(WhsCode, 1) = '" + Gbn + "' And RIGHT(a.WhsCode, 2) = RIGHT(b.DfltsGroup, 2)";
//			oRecordset01.DoQuery(sQry);

//			functionReturnValue = Strings.Trim(oRecordset01.Fields.Item(0).Value);
//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			return functionReturnValue;
//			User_WhsCode_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			functionReturnValue = Convert.ToString(0);
//			MDC_Com.MDC_GF_Message(ref "User_WhsCode_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			return functionReturnValue;
//		}

//		// 아이디별 사번 선택
//		public static string User_MSTCOD()
//		{
//			string functionReturnValue = null;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordset01 = null;

//			oRecordset01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "Select U_MSTCOD From [OHEM] a Inner Join [OUSR] b On a.userId = b.USERID Where b.USER_CODE = '" + Strings.Trim(SubMain.Sbo_Company.UserName) + "'";
//			oRecordset01.DoQuery(sQry);

//			functionReturnValue = Strings.Trim(oRecordset01.Fields.Item(0).Value);

//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			return functionReturnValue;
//			User_MSTCOD_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			functionReturnValue = Convert.ToString(0);
//			MDC_Com.MDC_GF_Message(ref "User_MSTCOD_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			return functionReturnValue;
//		}

//		// 아이디별 부서 선택
//		public static string User_DeptCode()
//		{
//			string functionReturnValue = null;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordset01 = null;

//			oRecordset01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "Select dept From [OHEM] a Inner Join [OUSR] b On a.userId = b.USERID Where USER_CODE = '" + Strings.Trim(SubMain.Sbo_Company.UserName) + "'";
//			oRecordset01.DoQuery(sQry);

//			functionReturnValue = Strings.Trim(oRecordset01.Fields.Item(0).Value);
//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			return functionReturnValue;
//			User_DeptCode_Error:
//			//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			functionReturnValue = Convert.ToString(0);
//			MDC_Com.MDC_GF_Message(ref "User_DeptCode_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			return functionReturnValue;
//		}

//		public static string User_TeamCode()
//		{
//			string functionReturnValue = null;
//			//******************************************************************************
//			//Function ID : User_TeamCode()
//			//해당모듈    : MDC_PS_Common
//			//기    능    : 접속한 사용자의 팀코드 조회
//			//인    수    : 없음
//			//반 환 값    : TeamCode
//			//특이사항    : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordset01 = null;

//			oRecordset01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "Select U_TeamCode From [OHEM] a Inner Join [OUSR] b On a.userId = b.USERID Where USER_CODE = '" + Strings.Trim(SubMain.Sbo_Company.UserName) + "'";
//			oRecordset01.DoQuery(sQry);

//			functionReturnValue = Strings.Trim(oRecordset01.Fields.Item(0).Value);
//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			return functionReturnValue;
//			User_TeamCode_Error:

//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			functionReturnValue = Convert.ToString(0);
//			MDC_Com.MDC_GF_Message(ref "User_TeamCode_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			return functionReturnValue;
//		}

//		public static string User_RspCode()
//		{
//			string functionReturnValue = null;
//			//******************************************************************************
//			//Function ID : User_RspCode()
//			//해당모듈    : MDC_PS_Common
//			//기    능    : 접속한 사용자의 담당코드 조회
//			//인    수    : 없음
//			//반 환 값    : RspCode
//			//특이사항    : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordset01 = null;

//			oRecordset01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "Select U_RspCode From [OHEM] a Inner Join [OUSR] b On a.userId = b.USERID Where USER_CODE = '" + Strings.Trim(SubMain.Sbo_Company.UserName) + "'";
//			oRecordset01.DoQuery(sQry);

//			functionReturnValue = Strings.Trim(oRecordset01.Fields.Item(0).Value);
//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			return functionReturnValue;
//			User_RspCode_Error:

//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			functionReturnValue = Convert.ToString(0);
//			MDC_Com.MDC_GF_Message(ref "User_RspCode_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			return functionReturnValue;
//		}

//		public static string User_ClsCode()
//		{
//			string functionReturnValue = null;
//			//******************************************************************************
//			//Function ID : User_ClsCode()
//			//해당모듈    : MDC_PS_Common
//			//기    능    : 접속한 사용자의 반코드 조회
//			//인    수    : 없음
//			//반 환 값    : ClsCode
//			//특이사항    : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordset01 = null;

//			oRecordset01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "Select U_ClsCode From [OHEM] a Inner Join [OUSR] b On a.userId = b.USERID Where USER_CODE = '" + Strings.Trim(SubMain.Sbo_Company.UserName) + "'";
//			oRecordset01.DoQuery(sQry);

//			functionReturnValue = Strings.Trim(oRecordset01.Fields.Item(0).Value);
//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			return functionReturnValue;
//			User_ClsCode_Error:

//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			functionReturnValue = Convert.ToString(0);
//			MDC_Com.MDC_GF_Message(ref "User_ClsCode_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			return functionReturnValue;
//		}

//		public static string User_SuperUserYN()
//		{
//			string functionReturnValue = null;
//			//******************************************************************************
//			//Function ID : User_SuperUserYN()
//			//해당모듈    : MDC_PS_Common
//			//기    능    : 접속한 사용자의 SuperUser 여부
//			//인    수    : 없음
//			//반 환 값    : Y:수퍼유저, N:일반유저
//			//특이사항    : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordset01 = null;

//			oRecordset01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "           SELECT      T0.SUPERUSER";
//			sQry = sQry + " FROM       OUSR AS T0";
//			sQry = sQry + " WHERE      T0.User_Code = '" + Strings.Trim(SubMain.Sbo_Company.UserName) + "'";

//			oRecordset01.DoQuery(sQry);

//			functionReturnValue = Strings.Trim(oRecordset01.Fields.Item(0).Value);
//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			return functionReturnValue;
//			User_SuperUserYN_Error:

//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			functionReturnValue = Convert.ToString(0);
//			MDC_Com.MDC_GF_Message(ref "User_SuperUserYN_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			return functionReturnValue;
//		}

//		public static string Future_Date_Check(string inputdate)
//		{
//			string functionReturnValue = null;
//			//******************************************************************************
//			//Function ID : Future_Date_Check()
//			//해당모듈    : MDC_PS_Common
//			//기    능    : 입력일자가 현재 서버일자보다 미래될수 없도록 제한함.
//			//인    수    :
//			//반 환 값    :
//			//특이사항    : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordset01 = null;

//			oRecordset01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "select case when convert(char(8),getdate(),112) >= '" + Strings.Trim(inputdate) + "'";
//			sQry = sQry + " then 'Y' else 'N' end";

//			oRecordset01.DoQuery(sQry);

//			functionReturnValue = Strings.Trim(oRecordset01.Fields.Item(0).Value);
//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			return functionReturnValue;
//			Future_Date_Check_Error:

//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			functionReturnValue = Convert.ToString(0);
//			MDC_Com.MDC_GF_Message(ref "User_SuperUserYN_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			return functionReturnValue;
//		}

//		public static string User_MainJob()
//		{
//			string functionReturnValue = null;
//			//******************************************************************************
//			//Function ID : User_MainJob()
//			//해당모듈    : MDC_PS_Common
//			//기    능    : 접속한 사용자의 주요업무 조회
//			//인    수    : 없음
//			//반 환 값    : 주요업무(인사마스터(OHEM)의 Remark)
//			//특이사항    : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordset01 = null;

//			oRecordset01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "           SELECT       T0.Remark";
//			sQry = sQry + " FROM        OHEM AS T0";
//			sQry = sQry + "                 LEFT JOIN";
//			sQry = sQry + "                 OUSR AS T1";
//			sQry = sQry + "                     ON T0.UserID = T1.USERID";
//			sQry = sQry + " WHERE       T1.User_Code = '" + Strings.Trim(SubMain.Sbo_Company.UserName) + "'";

//			oRecordset01.DoQuery(sQry);

//			functionReturnValue = Strings.Trim(oRecordset01.Fields.Item(0).Value);
//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			return functionReturnValue;
//			User_MainJob_Error:

//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			functionReturnValue = "";
//			MDC_Com.MDC_GF_Message(ref "User_MainJob_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			return functionReturnValue;
//		}

//		public static double Calculate_Weight(string ItemCode, int Qty, string BPLId)
//		{
//			double functionReturnValue = 0;
//			 // ERROR: Not supported in C#: OnErrorStatement


//			double ReturnValue = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordset01 = null;

//			oRecordset01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "Select U_OBasUnit, U_UnitQ1, U_Spec1, U_Spec2, U_Spec3, U_UnWeight From [OITM] Where ItemCode = '" + ItemCode + "'";
//			oRecordset01.DoQuery(sQry);

//			if (Strings.Trim(oRecordset01.Fields.Item(0).Value) == "101") {
//				ReturnValue = Qty;
//			} else if (Strings.Trim(oRecordset01.Fields.Item(0).Value) == "102") {
//				ReturnValue = Qty * Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(1).Value));
//			} else if (Strings.Trim(oRecordset01.Fields.Item(0).Value) == "201") {
//				ReturnValue = (Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(2).Value)) - Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(3).Value))) * Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(3).Value)) * 0.02808 * (Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(4).Value)) / 1000) * Qty;
//			} else if (Strings.Trim(oRecordset01.Fields.Item(0).Value) == "202") {
//				ReturnValue = Qty * Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(5).Value)) / 1000;
//			} else if (Strings.Trim(oRecordset01.Fields.Item(0).Value) == "203") {
//				ReturnValue = 0;
//			}

//			if (BPLId == "3" | BPLId == "5") {
//				functionReturnValue = System.Math.Round(ReturnValue, 2);
//			} else {
//				functionReturnValue = System.Math.Round(ReturnValue, 0);
//			}

//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			return functionReturnValue;
//			Calculate_Weight_Error:

//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			functionReturnValue = 0;
//			MDC_Com.MDC_GF_Message(ref "Calculate_Weight_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			return functionReturnValue;
//		}

//		public static int Calculate_Qty(string ItemCode, int Weight)
//		{
//			int functionReturnValue = 0;
//			 // ERROR: Not supported in C#: OnErrorStatement


//			double ReturnValue = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordset01 = null;

//			oRecordset01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "Select U_OBasUnit, U_UnitQ1, U_Spec1, U_Spec2, U_Spec3, U_UnWeight From [OITM] Where ItemCode = '" + ItemCode + "'";
//			oRecordset01.DoQuery(sQry);

//			if (Strings.Trim(oRecordset01.Fields.Item(0).Value) == "101") {
//				ReturnValue = Weight;
//			} else if (Strings.Trim(oRecordset01.Fields.Item(0).Value) == "102") {
//				if (string.IsNullOrEmpty(Strings.Trim(oRecordset01.Fields.Item(1).Value)) | Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(1).Value)) == 0) {
//					ReturnValue = 0;
//				} else {
//					ReturnValue = Weight / Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(1).Value));
//				}
//			} else if (Strings.Trim(oRecordset01.Fields.Item(0).Value) == "201") {
//				if ((Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(2).Value)) - Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(3).Value))) * Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(3).Value)) * 0.02808 * (Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(4).Value)) / 1000) == Convert.ToDouble("") | (Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(2).Value)) - Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(3).Value))) * Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(3).Value)) * 0.02808 * (Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(4).Value)) / 1000) == 0) {
//					ReturnValue = 0;
//				} else {
//					ReturnValue = Weight / ((Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(2).Value)) - Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(3).Value))) * Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(3).Value)) * 0.02808 * (Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(4).Value)) / 1000));
//				}
//			} else if (Strings.Trim(oRecordset01.Fields.Item(0).Value) == "202") {
//				if (string.IsNullOrEmpty(Strings.Trim(oRecordset01.Fields.Item(5).Value)) | Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(5).Value)) == 0) {
//					ReturnValue = 0;
//				} else {
//					ReturnValue = Weight / Convert.ToDouble(Strings.Trim(oRecordset01.Fields.Item(5).Value)) * 1000;
//				}
//			} else if (Strings.Trim(oRecordset01.Fields.Item(0).Value) == "203") {
//				ReturnValue = 0;
//			}

//			functionReturnValue = System.Math.Round(ReturnValue, 0);
//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			return functionReturnValue;
//			Calculate_Qty_Error:

//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			functionReturnValue = 0;
//			MDC_Com.MDC_GF_Message(ref "Calculate_Qty_Error:" + Err().Number + " - " + Err().Description, ref "E");
//			return functionReturnValue;
//		}

//		public static string RFC_Sender(string BPLId, string ItemCode, string ItemName, string Size, double Qty, string Unit, string RequestDate, string DueDate, string ItemType, string RequestNo,
//		int i, int LastRow)
//		{
//			string functionReturnValue = null;
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string ReturnValue = null;
//			string WERKS = null;

//			if (i == 0) {
//				MDC_Globals.oSapConnection01 = Interaction.CreateObject("SAP.Functions");
//				//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				MDC_Globals.oSapConnection01.Connection.User = "ifuser";
//				//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				MDC_Globals.oSapConnection01.Connection.Password = "pdauser";
//				//        oSapConnection01.Connection.client = "710"
//				//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				MDC_Globals.oSapConnection01.Connection.Client = "210";
//				//        oSapConnection01.Connection.ApplicationServer = "192.1.11.7"
//				//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				MDC_Globals.oSapConnection01.Connection.ApplicationServer = "192.1.1.217";
//				//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				MDC_Globals.oSapConnection01.Connection.Language = "KO";
//				//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				MDC_Globals.oSapConnection01.Connection.SystemNumber = "00";
//				//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (!MDC_Globals.oSapConnection01.Connection.Logon(0, true)) {
//					MDC_Com.MDC_GF_Message(ref "안강(R/3)서버에 접속할수 없습니다.", ref "E");
//					goto RFC_Sender_Exit;
//				}
//			}

//			object oFunction01 = null;
//			//UPGRADE_WARNING: oSapConnection01.Add 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oFunction01 = MDC_Globals.oSapConnection01.Add("ZMM_INTF_GROUP");
//			if (Convert.ToDouble(BPLId) == 1) {
//				WERKS = "9200";
//			} else if (Convert.ToDouble(BPLId) == 2) {
//				WERKS = "9300";
//			} else {
//				WERKS = "9200";
//			}

//			//UPGRADE_WARNING: oFunction01.Exports 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oFunction01.Exports("I_WERKS") = WERKS;
//			////플랜트 홀딩스 창원 9200, 홀딩스 부산 9300
//			//UPGRADE_WARNING: oFunction01.Exports 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oFunction01.Exports("I_MATNR") = ItemCode;
//			////자재코드 char(18)
//			//UPGRADE_WARNING: oFunction01.Exports 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oFunction01.Exports("I_MAKTX") = ItemName;
//			////자재내역 char(40)
//			//UPGRADE_WARNING: oFunction01.Exports 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oFunction01.Exports("I_WRKST") = Size;
//			////자재규격 char(48)
//			//UPGRADE_WARNING: oFunction01.Exports 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oFunction01.Exports("I_MENGE") = Qty;
//			////구매요청수량 dec(13,3)
//			//UPGRADE_WARNING: oFunction01.Exports 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oFunction01.Exports("I_MEINS") = Unit;
//			////단위 char(3)
//			//UPGRADE_WARNING: oFunction01.Exports 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oFunction01.Exports("I_BADAT") = RequestDate;
//			////구매요청일 char(8)
//			//UPGRADE_WARNING: oFunction01.Exports 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oFunction01.Exports("I_LFDAT") = DueDate;
//			////납품일 char(8)
//			//UPGRADE_WARNING: oFunction01.Exports 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oFunction01.Exports("I_MATKL") = ItemType;
//			////자재그룹 char(9)
//			//UPGRADE_WARNING: oFunction01.Exports 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oFunction01.Exports("I_ZBANFN") = RequestNo;
//			////구매요청번호

//			//UPGRADE_WARNING: oFunction01.Call 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (!(oFunction01.Call)) {
//				MDC_Com.MDC_GF_Message(ref "안강(R/3)서버 함수호출중 오류발생", ref "E");
//				goto RFC_Sender_Exit;
//			} else {
//				//UPGRADE_WARNING: oFunction01.Imports 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				////에러메시지
//				if ((string.IsNullOrEmpty(oFunction01.Imports("E_MESSAGE").VALUE))) {
//					//UPGRADE_WARNING: oFunction01.Imports 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					ReturnValue = oFunction01.Imports("E_BANFN").VALUE + "/" + oFunction01.Imports("E_BNFPO").VALUE;
//					////통합구매요청번호 '//통합구매요청 품목번호
//				} else {
//					//UPGRADE_WARNING: oFunction01.Imports 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					MDC_Com.MDC_GF_Message(ref oFunction01.Imports("E_MESSAGE").VALUE, ref "E");
//					goto RFC_Sender_Exit;
//				}
//			}

//			//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if ((MDC_Globals.oSapConnection01.Connection != null)) {
//				if (i == LastRow) {
//					//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					MDC_Globals.oSapConnection01.Connection.Logoff();
//					//UPGRADE_NOTE: oSapConnection01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//					MDC_Globals.oSapConnection01 = null;
//				}
//			}

//			functionReturnValue = ReturnValue;
//			//UPGRADE_NOTE: oFunction01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oFunction01 = null;
//			return functionReturnValue;
//			RFC_Sender_Exit:
//			//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if ((MDC_Globals.oSapConnection01.Connection != null)) {
//				if (i == LastRow) {
//					//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					MDC_Globals.oSapConnection01.Connection.Logoff();
//					//UPGRADE_NOTE: oSapConnection01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//					MDC_Globals.oSapConnection01 = null;
//				}
//			}
//			functionReturnValue = "";
//			//UPGRADE_NOTE: oFunction01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oFunction01 = null;
//			return functionReturnValue;
//			RFC_Sender_Error:
//			//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if ((MDC_Globals.oSapConnection01.Connection != null)) {
//				if (i == LastRow) {
//					//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					MDC_Globals.oSapConnection01.Connection.Logoff();
//					//UPGRADE_NOTE: oSapConnection01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//					MDC_Globals.oSapConnection01 = null;
//				}
//			}
//			functionReturnValue = "";
//			//UPGRADE_NOTE: oFunction01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oFunction01 = null;
//			SubMain.Sbo_Application.SetStatusBarMessage("RFC_Sender_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		public static string Cal_KPI_Grade(short prmBaseEntry, short prmBaseLine, string prmTableName, string prmResult, string prmMonth)
//		{
//			string functionReturnValue = null;
//			//******************************************************************************
//			//Function ID : Cal_KPI_Grade()
//			//해당모듈    : MDC_PS_Common
//			//기    능    : KPI 평가등급 계산
//			//인    수    : prmBaseEntry(KPI목표문서번호), prmBaseLine(KPI목표문서행번호), prmTableName(KPI목표 테이블 명), prmResult(실적), prmMonth(실적등록월)
//			//반 환 값    : KPI평가등급
//			//특이사항    : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			//1. 해당KPI목표 테이블의 문서번호와 행번호를 이용하여 A~E까지의 값 조회
//			//2. 등급기준(최대, 최소)에 따라 분기문이 달라져야 하므로 등급기준이 최대인지, 최소인지 함께 조회

//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordset01 = null;
//			oRecordset01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "EXEC PS_Z_GetKPIGrade " + prmBaseEntry + "," + prmBaseLine + ",'" + prmTableName + "','" + prmResult + "', '" + prmMonth + "'";

//			oRecordset01.DoQuery(sQry);

//			//UPGRADE_WARNING: oRecordset01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			functionReturnValue = oRecordset01.Fields.Item("Grade").Value;

//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			return functionReturnValue;
//			Cal_KPI_Grade_Error:


//			functionReturnValue = "";
//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			SubMain.Sbo_Application.SetStatusBarMessage("Cal_KPI_Grade_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;

//		}


//		public static double Cal_KPI_Score(string prmKPIGrade)
//		{
//			double functionReturnValue = 0;
//			//******************************************************************************
//			//Function ID : Cal_KPI_Score()
//			//해당모듈    : MDC_PS_Common
//			//기    능    : KPI 평가점수 계산
//			//인    수    : prmKPIGrade(KPI평가등급)
//			//반 환 값    : KPI평가점수
//			//특이사항    : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string sQry = null;
//			double KPI_Score = 0;

//			short loopCount01 = 0;

//			SAPbobsCOM.Recordset oRecordset01 = null;
//			oRecordset01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "        SELECT      T1.U_CodeNm AS [CodeName],";
//			sQry = sQry + "             T1.U_Num1 AS [KPIScore]";
//			sQry = sQry + " FROM        [@PS_HR200H] AS T0";
//			sQry = sQry + "             INNER JOIN";
//			sQry = sQry + "             [@PS_HR200L] AS T1";
//			sQry = sQry + "                 ON T0.Code = T1.Code";
//			sQry = sQry + " WHERE       T0.Name = '평가점수'";

//			oRecordset01.DoQuery(sQry);

//			for (loopCount01 = 0; loopCount01 <= oRecordset01.RecordCount - 1; loopCount01++) {

//				if (prmKPIGrade == oRecordset01.Fields.Item("CodeName").Value) {

//					//UPGRADE_WARNING: oRecordset01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					KPI_Score = oRecordset01.Fields.Item("KPIScore").Value;

//				}

//				oRecordset01.MoveNext();

//			}

//			functionReturnValue = KPI_Score;

//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			return functionReturnValue;
//			Cal_KPI_Score_Error:


//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			SubMain.Sbo_Application.SetStatusBarMessage("Cal_KPI_Score_Error " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;

//		}

//		public static double Cal_KPI_AchieveRate(short prmBasEntry, short prmBasLine, string prmDocType, string prmRslt)
//		{
//			double functionReturnValue = 0;
//			//******************************************************************************
//			//Function ID : Cal_KPI_AchieveRate()
//			//해당모듈    : MDC_PS_Common
//			//기    능    : KPI 진척율(달성율)
//			//인    수    : prmBasEntry(목표문서번호), prmBasLine(목표행번호), prmDocType(문서타입), prmRslt(실적)
//			//반 환 값    : KPI평가점수
//			//특이사항    : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordset01 = null;
//			oRecordset01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			sQry = "EXEC PS_Z_GetAchieveRate " + prmBasEntry + "," + prmBasLine + ",'" + prmDocType + "','" + prmRslt + "'";
//			//진척율 계산 프로시져

//			oRecordset01.DoQuery(sQry);

//			//UPGRADE_WARNING: oRecordset01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			functionReturnValue = oRecordset01.Fields.Item("AchieveRate").Value;

//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			return functionReturnValue;
//			Cal_KPI_AchieveRate_Error:


//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			SubMain.Sbo_Application.SetStatusBarMessage("Cal_KPI_AchieveRate_Error " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;

//		}

//		public static bool Check_Finish_Status(string prmBPLId, string prmDocDate, object prmFormTypeEx)
//		{
//			bool functionReturnValue = false;
//			//******************************************************************************
//			//Function ID : Check_Finish_Status()
//			//해당모듈    : MDC_PS_Common
//			//기    능    : 마감상태 조회
//			//인    수    : prmBPLID(사업장), prmDocDate(등록일), prmFormTypeEx(화면타입(UID))
//			//반 환 값    : 마감상태에 따른 등록 가능 여부
//			//특이사항    : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordset01 = null;
//			oRecordset01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			string CheckFinishStatus = null;

//			sQry = "      EXEC PS_Z_CheckFinishStatus '";
//			sQry = sQry + prmBPLId + "','";
//			sQry = sQry + prmDocDate + "','";
//			//UPGRADE_WARNING: prmFormTypeEx 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = sQry + prmFormTypeEx + "'";

//			oRecordset01.DoQuery(sQry);

//			//UPGRADE_WARNING: oRecordset01.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CheckFinishStatus = oRecordset01.Fields.Item("ReturnValue").Value;

//			if (CheckFinishStatus == "True") {
//				functionReturnValue = true;
//			} else {
//				functionReturnValue = false;
//			}

//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			return functionReturnValue;
//			Check_Finish_Status_Error:


//			//UPGRADE_NOTE: oRecordset01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordset01 = null;
//			functionReturnValue = false;
//			SubMain.Sbo_Application.SetStatusBarMessage("Check_Finish_Status_Error " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		public static string Split_String(string pSplitString, string pSeparateChar, short pIndex)
//		{
//			string functionReturnValue = null;
//			//******************************************************************************
//			//Function ID : Split_String()
//			//해당모듈    : MDC_PS_Common
//			//기    능    : 문자열 Split
//			//인    수    : pSplitString(대상 문자열), pSeparateChar(분할 기준 Char), pIndex(분할된 문자열 중 반환할 문자열의 Index)
//			//반 환 값    : 분할된 문자열
//			//특이사항    : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			object StringTemp = null;

//			//UPGRADE_WARNING: StringTemp 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			StringTemp = SAPbobsCOM.GTSResponseToExceedingEnum.Split(pSplitString, pSeparateChar);

//			if (pIndex > 0 & pIndex - 1 <= Information.UBound(StringTemp)) {
//				//UPGRADE_WARNING: StringTemp() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				functionReturnValue = StringTemp(pIndex - 1);
//			} else {
//				functionReturnValue = "";
//			}
//			return functionReturnValue;
//			Split_String_Error:



//			functionReturnValue = "";
//			SubMain.Sbo_Application.SetStatusBarMessage("Split_String_Error " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;

//		}
//	}
//}
