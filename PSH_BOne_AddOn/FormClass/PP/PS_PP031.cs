using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 품목별공수조회
	/// </summary>
	internal class PS_PP031 : PSH_BaseClass
	{
		public string oFormUniqueID;
		public SAPbouiCOM.Matrix oMat01;
		public SAPbouiCOM.Matrix oMat02;
		public SAPbouiCOM.Matrix oMat03;
		public SAPbouiCOM.Matrix oMat04;
		private SAPbouiCOM.DBDataSource oDS_PS_PP031L; //라인(품목분류별규격정보)
		private SAPbouiCOM.DBDataSource oDS_PS_PP031M; //라인(작업지시정보)
		private SAPbouiCOM.DBDataSource oDS_PS_PP031N; //라인(실동공수정보)
		private SAPbouiCOM.DBDataSource oDS_PS_PP031O; //라인(작업일보정보)
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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP031.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP031_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP031");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				
				oForm.Freeze(true);
                PS_PP031_CreateItems();
                PS_PP031_ComboBox_Setting();
                PS_PP031_FormResize();

                oForm.Items.Item("Folder01").Specific.Select();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_PP031_CreateItems()
        {
            try
            {
                oDS_PS_PP031L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oDS_PS_PP031M = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");
                oDS_PS_PP031N = oForm.DataSources.DBDataSources.Item("@PS_USERDS03");
                oDS_PS_PP031O = oForm.DataSources.DBDataSources.Item("@PS_USERDS04");

                //매트릭스 초기화
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oMat02 = oForm.Items.Item("Mat02").Specific;
                oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat02.AutoResizeColumns();

                oMat03 = oForm.Items.Item("Mat03").Specific;
                oMat03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat03.AutoResizeColumns();

                oMat04 = oForm.Items.Item("Mat04").Specific;
                oMat04.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat04.AutoResizeColumns();

                //품목대분류
                oForm.DataSources.UserDataSources.Add("ItmBsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ItmBsort").Specific.DataBind.SetBound(true, "", "ItmBsort");
                
                //품목구분
                oForm.DataSources.UserDataSources.Add("ItemClass", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ItemClass").Specific.DataBind.SetBound(true, "", "ItemClass");
                
                //거래처구분
                oForm.DataSources.UserDataSources.Add("CardType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("CardType").Specific.DataBind.SetBound(true, "", "CardType");
                
                //품명
                oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");
                
                //규격_S
                oForm.DataSources.UserDataSources.Add("Spec", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("Spec").Specific.DataBind.SetBound(true, "", "Spec");
                
                oMat03.Columns.Item("ItemCode").Visible = false; //실동공수정보 Matrix의 품목코드 필드 Hidden
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP031_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                //품목대분류
                oForm.Items.Item("ItmBsort").Specific.ValidValues.Add("", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("ItmBsort").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' AND Code IN ('105','106') order by Code", "", false, false);
                oForm.Items.Item("ItmBsort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //품목구분
                oForm.Items.Item("ItemClass").Specific.ValidValues.Add("", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("ItemClass").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'S002' ORDER BY Code", "", false, false);
                oForm.Items.Item("ItemClass").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //거래처구분
                oForm.Items.Item("CardType").Specific.ValidValues.Add("", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("CardType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'C100' ORDER BY Code", "", false, false);
                oForm.Items.Item("CardType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PS_PP031_FormResize()
        {
            try
            {
                oForm.Freeze(true);

                oForm.Items.Item("Mat01").Width = oForm.Width / 2 - 20;
                oForm.Items.Item("Mat01").Height = oForm.Height - oForm.Items.Item("Mat01").Top - (oForm.Height - oForm.Items.Item("2").Top) - 10;

                oForm.Items.Item("Mat02").Left = oForm.Width / 2 + 10;
                oForm.Items.Item("Mat02").Width = oForm.Width / 2 - 40;
                oForm.Items.Item("Mat02").Height = oForm.Items.Item("Mat01").Height - 20;

                oForm.Items.Item("Mat03").Left = oForm.Width / 2 + 10;
                oForm.Items.Item("Mat03").Width = oForm.Width / 2 - 40;
                oForm.Items.Item("Mat03").Height = oForm.Items.Item("Mat01").Height - 20;

                oForm.Items.Item("Mat04").Left = oForm.Width / 2 + 10;
                oForm.Items.Item("Mat04").Width = oForm.Width / 2 - 40;
                oForm.Items.Item("Mat04").Height = oForm.Items.Item("Mat01").Height - 20;

                oForm.Items.Item("Rec01").Left = oForm.Width / 2;
                oForm.Items.Item("Rec01").Height = oForm.Items.Item("Mat02").Height + 20;
                oForm.Items.Item("Rec01").Width = oForm.Items.Item("Mat02").Width + 20;

                oForm.Items.Item("Folder01").Width = 100;
                oForm.Items.Item("Folder01").Left = oForm.Items.Item("Rec01").Left;

                oForm.Items.Item("Folder02").Width = 100;
                oForm.Items.Item("Folder02").Left = oForm.Items.Item("Folder01").Left + oForm.Items.Item("Folder02").Width;

                oForm.Items.Item("Folder03").Width = 100;
                oForm.Items.Item("Folder03").Left = oForm.Items.Item("Folder02").Left + oForm.Items.Item("Folder03").Width;

                oMat01.AutoResizeColumns();
                oMat02.AutoResizeColumns();
                oMat03.AutoResizeColumns();
                oMat04.AutoResizeColumns();
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
        /// 메트릭스 Row추가(Mat01)
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param> 
        private void PS_PP031_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                if (RowIserted == false)
                {
                    oDS_PS_PP031L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_PP031L.Offset = oRow;
                oMat01.LoadFromDataSource();
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
        /// 메트릭스 데이터 로드
        /// </summary>
        private void PS_PP031_MTX01()
        {
            
            int loopCount;
            string Query01;
            string ItmBsort; //품목대분류
            string CardType; //거래처구분
            string ItemName; //품명
            string ItemClass; //품목구분
            string SPEC; //규격
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = null;

            try
            {
                oForm.Freeze(true);

                ItmBsort = oForm.Items.Item("ItmBsort").Specific.Selected.Value;
                CardType = oForm.Items.Item("CardType").Specific.Selected.Value.ToString().Trim();
                ItemClass = oForm.Items.Item("ItemClass").Specific.Selected.Value.ToString().Trim();
                ItemName = oForm.Items.Item("ItemName").Specific.Value;
                SPEC = oForm.Items.Item("Spec").Specific.Value;

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                Query01 = "EXEC PS_PP031_01 '" + ItmBsort + "','" + CardType + "','" + ItemClass + "','" + ItemName + "','" + SPEC + "'";
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

                for (loopCount = 0; loopCount <= RecordSet01.RecordCount - 1; loopCount++)
                {
                    if (loopCount != 0)
                    {
                        oDS_PS_PP031L.InsertRecord(loopCount);
                    }
                    oDS_PS_PP031L.Offset = loopCount;

                    oDS_PS_PP031L.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1)); //라인번호
                    oDS_PS_PP031L.SetValue("U_ColReg01", loopCount, RecordSet01.Fields.Item("ItmBName").Value); //품목대분류명
                    oDS_PS_PP031L.SetValue("U_ColReg02", loopCount, RecordSet01.Fields.Item("ItmMName").Value); //품목구분
                    oDS_PS_PP031L.SetValue("U_ColReg03", loopCount, RecordSet01.Fields.Item("ItemName").Value); //품목명
                    oDS_PS_PP031L.SetValue("U_ColReg04", loopCount, RecordSet01.Fields.Item("Spec").Value); //규격
                    oDS_PS_PP031L.SetValue("U_ColReg05", loopCount, RecordSet01.Fields.Item("Cnt").Value); //Count

                    RecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(errMessage, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                if (ProgressBar01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
            }
        }

        /// <summary>
        /// 메트릭스 데이터 로드
        /// </summary>
        /// <param name="prmItemName"></param>
        /// <param name="prmSpec"></param>
        private void PS_PP031_MTX02(string prmItemName, string prmSpec)
        {
            int loopCount;
            string Query01;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = null;

            try
            {
                oForm.Freeze(true);

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, false);

                Query01 = "EXEC PS_PP031_02 '" + prmItemName + "','" + prmSpec + "'";
                RecordSet01.DoQuery(Query01);

                oMat02.Clear();
                oMat02.FlushToDataSource();
                oMat02.LoadFromDataSource();

                if (RecordSet01.RecordCount == 0)
                {
                    oMat02.Clear();
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }

                for (loopCount = 0; loopCount <= RecordSet01.RecordCount - 1; loopCount++)
                {
                    if (loopCount != 0)
                    {
                        oDS_PS_PP031M.InsertRecord(loopCount);
                    }
                    oDS_PS_PP031M.Offset = loopCount;

                    oDS_PS_PP031M.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1)); //라인번호
                    oDS_PS_PP031M.SetValue("U_ColReg01", loopCount, RecordSet01.Fields.Item("WOEntry").Value); //작업지시문서번호
                    oDS_PS_PP031M.SetValue("U_ColReg02", loopCount, RecordSet01.Fields.Item("ItemCode").Value); //품목코드
                    oDS_PS_PP031M.SetValue("U_ColReg03", loopCount, RecordSet01.Fields.Item("ItemName").Value); //품목명
                    oDS_PS_PP031M.SetValue("U_ColReg04", loopCount, RecordSet01.Fields.Item("Spec").Value); //규격
                    oDS_PS_PP031M.SetValue("U_ColReg05", loopCount, RecordSet01.Fields.Item("WODate").Value); //등록일자

                    RecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
                }
                oMat02.LoadFromDataSource();
                oMat02.AutoResizeColumns();
            }
            catch(Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(errMessage, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                if (ProgressBar01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
            }
        }

        /// <summary>
        /// 메트릭스 데이터 로드
        /// </summary>
        /// <param name="prmItemCode"></param>
        private void PS_PP031_MTX03(string prmItemCode)
        {
            int loopCount;
            string Query01;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = null;

            try
            {
                oForm.Freeze(true);

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
                
                Query01 = "EXEC PS_PP031_03 '" + prmItemCode + "'";
                RecordSet01.DoQuery(Query01);

                oMat03.Clear();
                oMat03.FlushToDataSource();
                oMat03.LoadFromDataSource();

                if (RecordSet01.RecordCount == 0)
                {
                    oMat03.Clear();
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }

                for (loopCount = 0; loopCount <= RecordSet01.RecordCount - 1; loopCount++)
                {
                    if (loopCount != 0)
                    {
                        oDS_PS_PP031N.InsertRecord(loopCount);
                    }
                    oDS_PS_PP031N.Offset = loopCount;

                    oDS_PS_PP031N.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1)); //라인번호
                    oDS_PS_PP031N.SetValue("U_ColReg01", loopCount, RecordSet01.Fields.Item("CpBCode").Value); //공정대분류
                    oDS_PS_PP031N.SetValue("U_ColReg02", loopCount, RecordSet01.Fields.Item("CpBName").Value); //공정대분류명
                    oDS_PS_PP031N.SetValue("U_ColReg03", loopCount, RecordSet01.Fields.Item("CpCode").Value); //공정중분류
                    oDS_PS_PP031N.SetValue("U_ColReg04", loopCount, RecordSet01.Fields.Item("CpName").Value); //공정중분류명
                    oDS_PS_PP031N.SetValue("U_ColReg05", loopCount, RecordSet01.Fields.Item("WkTime").Value); //실동공수
                    oDS_PS_PP031N.SetValue("U_ColReg06", loopCount, RecordSet01.Fields.Item("ItemCode").Value); //품목코드

                    RecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
                }
                oMat03.LoadFromDataSource();
                oMat03.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(errMessage, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                if (ProgressBar01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
            }
        }

        /// <summary>
        /// 메트릭스 데이터 로드
        /// </summary>
        /// <param name="prmItemCode"></param>
        /// <param name="prmCpCode"></param>
        private void PS_PP031_MTX04(string prmItemCode, string prmCpCode)
        {
            int loopCount;
            string Query01;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = null;

            try
            {
                oForm.Freeze(true);

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                Query01 = "EXEC PS_PP031_04 '" + prmItemCode + "','" + prmCpCode + "'";
                RecordSet01.DoQuery(Query01);

                oMat04.Clear();
                oMat04.FlushToDataSource();
                oMat04.LoadFromDataSource();

                if (RecordSet01.RecordCount == 0)
                {
                    oMat04.Clear();
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }

                for (loopCount = 0; loopCount <= RecordSet01.RecordCount - 1; loopCount++)
                {
                    if (loopCount != 0)
                    {
                        oDS_PS_PP031O.InsertRecord(loopCount);
                    }
                    oDS_PS_PP031O.Offset = loopCount;

                    oDS_PS_PP031O.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1)); //라인번호
                    oDS_PS_PP031O.SetValue("U_ColReg01", loopCount, RecordSet01.Fields.Item("WREntry").Value); //작업일보문서번호
                    oDS_PS_PP031O.SetValue("U_ColReg02", loopCount, RecordSet01.Fields.Item("ItemCode").Value); //품목코드
                    oDS_PS_PP031O.SetValue("U_ColReg03", loopCount, RecordSet01.Fields.Item("ItemName").Value); //품목명
                    oDS_PS_PP031O.SetValue("U_ColReg04", loopCount, RecordSet01.Fields.Item("Spec").Value); //규격
                    oDS_PS_PP031O.SetValue("U_ColReg05", loopCount, RecordSet01.Fields.Item("CpCode").Value); //공정코드
                    oDS_PS_PP031O.SetValue("U_ColReg06", loopCount, RecordSet01.Fields.Item("CpName").Value); //공정명
                    oDS_PS_PP031O.SetValue("U_ColReg07", loopCount, RecordSet01.Fields.Item("WkTime").Value); //실동공수
                    oDS_PS_PP031O.SetValue("U_ColReg08", loopCount, RecordSet01.Fields.Item("WRDate").Value); //등록일자

                    RecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
                }
                oMat04.LoadFromDataSource();
                oMat04.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(errMessage, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                if (ProgressBar01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
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
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    //Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "BtnSearch")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            oMat02.Clear();
                            oDS_PS_PP031M.Clear();
                            PS_PP031_MTX01();
                        }
                    }
                    //else if (pVal.ItemUID == "Link04")
                    //{
                    //    PS_PP030 PP030 = new PS_PP030();
                    //    PP030.LoadForm(oForm.Items.Item("WODocNum").Specific.Value);
                    //}
                }
                else if (pVal.BeforeAction == false)
                {
                    //폴더를 사용할 때는 필수 소스_S
                    if (pVal.ItemUID == "Folder01") //Folder01이 선택되었을 때
                    {
                        oForm.PaneLevel = 1;
                    }
                    else if (pVal.ItemUID == "Folder02") //Folder02가 선택되었을 때
                    {
                        oForm.PaneLevel = 2;
                    }
                    else if (pVal.ItemUID == "Folder03") //Folder03이 선택되었을 때
                    {
                        oForm.PaneLevel = 3;
                    }
                    //폴더를 사용할 때는 필수 소스_E
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
        
            try
            {
                if (pVal.Before_Action == true)
                {
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", ""); //거래처코드
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", ""); //품목코드(작번)
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
                    else if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                    }
                    else if (pVal.ItemUID == "Mat03")
                    {
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                    }
                    else if (pVal.ItemUID == "Mat04")
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

                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                    }
                    else if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat02.SelectRow(pVal.Row, true, false);

                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                    }
                    else if (pVal.ItemUID == "Mat03")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat03.SelectRow(pVal.Row, true, false);

                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                    }
                    else if (pVal.ItemUID == "Mat04")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat04.SelectRow(pVal.Row, true, false);

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
                    if (pVal.ItemUID == "Mat01") //품목대분류별 규격 정보
                    {
                        if (pVal.Row == 0)
                        {   
                            oMat01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true; //정렬
                            oMat01.FlushToDataSource();
                        }
                        else
                        {
                            oForm.Items.Item("Folder01").Specific.Select(); //작업지시정보 TAB 선택
                            PS_PP031_MTX02(oMat01.Columns.Item("ItemName").Cells.Item(pVal.Row).Specific.Value, oMat01.Columns.Item("Spec").Cells.Item(pVal.Row).Specific.Value);
                        }
                    }
                    else if (pVal.ItemUID == "Mat02") //작업지시정보
                    {
                        if (pVal.Row == 0)
                        {   
                            oMat02.Columns.Item(pVal.ColUID).TitleObject.Sortable = true; //정렬
                            oMat02.FlushToDataSource();
                        }
                        else
                        {
                            oForm.Items.Item("Folder02").Specific.Select(); //실동공수정보 TAB 선택
                            PS_PP031_MTX03(oMat02.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value);
                        }
                    }
                    else if (pVal.ItemUID == "Mat03") //실동공수정보
                    {
                        if (pVal.Row == 0)
                        {
                            oMat03.Columns.Item(pVal.ColUID).TitleObject.Sortable = true; //정렬
                            oMat03.FlushToDataSource();
                        }
                        else
                        {
                            oForm.Items.Item("Folder03").Specific.Select(); //실동공수정보 TAB 선택
                            PS_PP031_MTX04(oMat03.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value, oMat03.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value);
                        }
                    }
                    else if (pVal.ItemUID == "Mat04") //작업일보정보
                    {
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
        /// MATRIX_LINK_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.ColUID == "WOEntry")
                        {
                            PS_PP030 PP030 = new PS_PP030();
                            PP030.LoadForm(oMat02.Columns.Item("WOEntry").Cells.Item(pVal.Row).Specific.Value);
                        }
                    }
                    else if (pVal.ItemUID == "Mat04")
                    {
                        if (pVal.ColUID == "WREntry")
                        {
                            PS_PP040 PP040 = new PS_PP040();
                            PP040.LoadForm(oMat04.Columns.Item("WREntry").Cells.Item(pVal.Row).Specific.Value);
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
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat03);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat04);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP031L);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP031M);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP031N);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP031O);
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
                    PS_PP031_FormResize();
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
                        case "7169": //엑셀 내보내기
                            //엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                            PS_PP031_AddMatrixRow(oMat01.VisualRowCount, false);
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
                            break;
                        case "1282": //추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                        case "7169": //엑셀 내보내기
                            //엑셀 내보내기 이후 처리
                            oForm.Freeze(true);
                            oDS_PS_PP031L.RemoveRecord(oDS_PS_PP031L.Size - 1);
                            oMat01.LoadFromDataSource();
                            oForm.Freeze(false);
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
                        Raise_EVENT_FORM_DATA_LOAD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                        Raise_EVENT_FORM_DATA_ADD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                        Raise_EVENT_FORM_DATA_UPDATE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                        Raise_EVENT_FORM_DATA_DELETE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
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
        /// FORM_DATA_LOAD 이벤트
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_FORM_DATA_LOAD(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FORM_DATA_ADD 이벤트
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_FORM_DATA_ADD(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FORM_DATA_UPDATE 이벤트
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_FORM_DATA_UPDATE(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FORM_DATA_DELETE 이벤트
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_FORM_DATA_DELETE(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
