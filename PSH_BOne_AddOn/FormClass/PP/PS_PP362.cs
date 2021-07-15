using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using System.Collections.Generic;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 5-1.작번별 생산진행현황2
    /// </summary>
    internal class PS_PP362 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Matrix oMat02;
        private SAPbouiCOM.Matrix oMat03;
        private SAPbouiCOM.Matrix oMat04;
        private SAPbouiCOM.Matrix oMat05;
        private SAPbouiCOM.Matrix oMat06;
        private SAPbouiCOM.Matrix oMat07;
        private SAPbouiCOM.Matrix oMat08;

        private SAPbouiCOM.DBDataSource oDS_PS_PP362L;        //라인(Sub작번)
        private SAPbouiCOM.DBDataSource oDS_PS_PP362M;        //라인(자재비내역)
        private SAPbouiCOM.DBDataSource oDS_PS_PP362N;        //라인(설계비내역)
        private SAPbouiCOM.DBDataSource oDS_PS_PP362O;        //라인(자체가공비내역_공정별내역)
        private SAPbouiCOM.DBDataSource oDS_PS_PP362P;        //라인(자체가공비내역_작업자별내역)
        private SAPbouiCOM.DBDataSource oDS_PS_PP362Q;       //라인(외주가공비내역)
        private SAPbouiCOM.DBDataSource oDS_PS_PP362R;        //라인(외주제작비내역)
        private SAPbouiCOM.DBDataSource oDS_PS_PP362S;
       
        private string oLastItemUID01;       //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01;        //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01;           //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP362.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP362_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP362");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PS_PP362_CreateItems();
                PS_PP362_ComboBox_Setting();
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
                oForm.Items.Item("Folder01").Specific.Select();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_PP362_CreateItems()
        {
            try
            {
                oDS_PS_PP362L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oDS_PS_PP362M = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");
                oDS_PS_PP362N = oForm.DataSources.DBDataSources.Item("@PS_USERDS03");
                oDS_PS_PP362O = oForm.DataSources.DBDataSources.Item("@PS_USERDS04");
                oDS_PS_PP362P = oForm.DataSources.DBDataSources.Item("@PS_USERDS05");
                oDS_PS_PP362Q = oForm.DataSources.DBDataSources.Item("@PS_USERDS06");
                oDS_PS_PP362R = oForm.DataSources.DBDataSources.Item("@PS_USERDS07");
                oDS_PS_PP362S = oForm.DataSources.DBDataSources.Item("@PS_USERDS08");

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

                oMat05 = oForm.Items.Item("Mat05").Specific;
                oMat05.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat05.AutoResizeColumns();

                oMat06 = oForm.Items.Item("Mat06").Specific;
                oMat06.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat06.AutoResizeColumns();

                oMat07 = oForm.Items.Item("Mat07").Specific;
                oMat07.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat07.AutoResizeColumns();

                oMat08 = oForm.Items.Item("Mat08").Specific;
                oMat08.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat08.AutoResizeColumns();

                //작번등록년월(시작)_S
                oForm.DataSources.UserDataSources.Add("FrDt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 7);
                oForm.Items.Item("FrDt").Specific.DataBind.SetBound(true, "", "FrDt");
                oForm.DataSources.UserDataSources.Item("FrDt").Value = DateTime.Now.ToString("yyyy-MM");
                //작번등록년월(시작)_E

                //작번등록년월(종료)_S
                oForm.DataSources.UserDataSources.Add("ToDt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 7);
                oForm.Items.Item("ToDt").Specific.DataBind.SetBound(true, "", "ToDt");
                oForm.DataSources.UserDataSources.Item("ToDt").Value = DateTime.Now.ToString("yyyy-MM");
                //작번등록년월(종료)_E

                //품명_S
                oForm.DataSources.UserDataSources.Add("FrgnName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("FrgnName").Specific.DataBind.SetBound(true, "", "FrgnName");
                //품명_E

                //거래처구분_S
                oForm.DataSources.UserDataSources.Add("CardType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CardType").Specific.DataBind.SetBound(true, "", "CardType");
                //거래처구분_E

                //규격_S
                oForm.DataSources.UserDataSources.Add("Spec", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("Spec").Specific.DataBind.SetBound(true, "", "Spec");
                //규격_E

                //품목구분_S
                oForm.DataSources.UserDataSources.Add("ItemClass", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ItemClass").Specific.DataBind.SetBound(true, "", "ItemClass");
                //품목구분_E

                //생산완료여부_S
                oForm.DataSources.UserDataSources.Add("WCYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("WCYN").Specific.DataBind.SetBound(true, "", "WCYN");
                //생산완료여부_E

                //거래처_S
                oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("CardCode").Specific.DataBind.SetBound(true, "", "CardCode");
                //거래처_E

                //거래처명_S
                oForm.DataSources.UserDataSources.Add("CardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("CardName").Specific.DataBind.SetBound(true, "", "CardName");
                //거래처명_E

                //품목(작번)_S
                oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");
                //품목(작번)_E

                //품목명_S
                oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");
                //품목명_E

                //Main작번_S
                oForm.DataSources.UserDataSources.Add("MainJak", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("MainJak").Specific.DataBind.SetBound(true, "", "MainJak");
                //Main작번_E

                //Sub작번_S
                oForm.DataSources.UserDataSources.Add("SubJak", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("SubJak").Specific.DataBind.SetBound(true, "", "SubJak");
                //Sub작번_E

                //공정코드_S
                oForm.DataSources.UserDataSources.Add("CpCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("CpCode").Specific.DataBind.SetBound(true, "", "CpCode");
                //공정코드_E

                //자체가공비내역_공정별내역 Matrix의 필드 Hidden
                oMat05.Columns.Item("ReTime").Visible = false;                //수정공수
                oMat05.Columns.Item("CpCode").Visible = false;                //공정코드
                oMat05.Columns.Item("OrdNum").Visible = false;                //작번
                oMat05.Columns.Item("Sub1_2").Visible = false;                //서브작번
                oMat05.Columns.Item("Class").Visible = false;                //공정완료여부(2012.08.06 송명규 수정)

                //자체가공비내역의 공정코드 필드 Hidden
                oForm.Items.Item("CpCode").Visible = false;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP362_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("CardType").Specific.ValidValues.Add("%", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("CardType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'C100' ORDER BY Code", "", false, false);
                oForm.Items.Item("CardType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                //거래처구분_E

                //품목구분_S
                oForm.Items.Item("ItemClass").Specific.ValidValues.Add("%", "전체");
                dataHelpClass.Set_ComboList(oForm.Items.Item("ItemClass").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'S002' ORDER BY Code", "", false, false);
                oForm.Items.Item("ItemClass").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                //품목구분_E

                //생산완료여부_S
                oForm.Items.Item("WCYN").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("WCYN").Specific.ValidValues.Add("B", "미완료");
                oForm.Items.Item("WCYN").Specific.ValidValues.Add("C", "완료");
                oForm.Items.Item("WCYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                //생산완료여부_E
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PS_PP362_FormResize()
        {
            try
            {
                oForm.Items.Item("Mat05").Width = oForm.Width / 2 - 25;
                oForm.Items.Item("Mat05").Height = oForm.Height - oForm.Items.Item("Mat05").Top - (oForm.Height - oForm.Items.Item("Btn05").Top) - 10;

                oForm.Items.Item("Mat06").Left = oForm.Width / 2 + 5;
                oForm.Items.Item("Mat06").Width = oForm.Width / 2 - 45;
                oForm.Items.Item("Mat06").Height = oForm.Items.Item("Mat05").Height;

                oForm.Items.Item("Static11").Left = oForm.Items.Item("Mat06").Left;

                oForm.Items.Item("Btn06").Top = oForm.Items.Item("Btn05").Top;
                oForm.Items.Item("Btn06").Left = oForm.Items.Item("Static11").Left;

                oForm.Items.Item("Rec01").Height = oForm.Items.Item("Mat01").Height + 50;
                oForm.Items.Item("Rec01").Width = oForm.Items.Item("Mat01").Width + 25;

                oMat01.AutoResizeColumns();
                oMat02.AutoResizeColumns();
                oMat03.AutoResizeColumns();
                oMat04.AutoResizeColumns();
                oMat05.AutoResizeColumns();
                oMat06.AutoResizeColumns();
                oMat07.AutoResizeColumns();
                oMat08.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PS_PP362_AddMatrixRow
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="prmMat"></param>
        /// <param name="prmDataSource"></param>
        /// <param name="RowIserted"></param>
        private void PS_PP362_AddMatrixRow(int oRow, SAPbouiCOM.Matrix prmMat, SAPbouiCOM.DBDataSource prmDataSource, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                //행추가여부
                if (RowIserted == false)
                {
                    prmDataSource.InsertRecord(oRow);
                }
                prmMat.AddRow();
                prmDataSource.Offset = oRow;
                //    oDS_PS_PP362L.setValue "U_LineNum", oRow, oRow + 1
                prmMat.LoadFromDataSource();
                oForm.Freeze(false);
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
        /// PS_PP362_MTX01
        /// </summary>
        private void PS_PP362_MTX01()
        {
            int loopCount;
            string Query01;
            string FrDt;      //작번등록년월(시작)
            string ToDt;      //작번등록년월(종료)
            string FrgnName;  //품명
            string CardType;  //거래처구분
            string SPEC;      //규격
            string ItemClass; //품목구분
            string WCYN;      //생산완료여부
            string CardCode;  //거래처
            string ItemCode;  //품목(작번)
            string errMessage = string.Empty;

            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
                ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();
                FrgnName = (string.IsNullOrEmpty(oForm.Items.Item("FrgnName").Specific.Value) ? "%" : oForm.Items.Item("FrgnName").Specific.Value);
                CardType = oForm.Items.Item("CardType").Specific.Selected.Value.ToString().Trim();
                SPEC = (string.IsNullOrEmpty(oForm.Items.Item("Spec").Specific.Value) ? "%" : oForm.Items.Item("Spec").Specific.Value);
                ItemClass = oForm.Items.Item("ItemClass").Specific.Selected.Value.ToString().Trim();
                WCYN = (string.IsNullOrEmpty(oForm.Items.Item("WCYN").Specific.Selected.Value) ? "%" : oForm.Items.Item("WCYN").Specific.Selected.Value);
                CardCode = (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value) ? "%" : oForm.Items.Item("CardCode").Specific.Value);
                ItemCode = (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value) ? "%" : oForm.Items.Item("ItemCode").Specific.Value);

                ProgressBar01.Text = "조회시작!"; //쿼리를 실행할 때 부터 프로그레스 시작
                oForm.Freeze(true);
                Query01 = "EXEC PS_PP362_01 '" + FrDt + "','" + ToDt + "','" + FrgnName + "','" + CardType + "','" + SPEC + "','" + ItemClass + "','" + WCYN + "','" + CardCode + "','" + ItemCode + "'";
                oRecordSet01.DoQuery(Query01);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    oMat01.Clear();
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }
                for (loopCount = 0; loopCount <= oRecordSet01.RecordCount - 1; loopCount++)
                {
                    if (loopCount != 0)
                    {
                        oDS_PS_PP362L.InsertRecord(loopCount);
                    }
                    oDS_PS_PP362L.Offset = loopCount;
                    oDS_PS_PP362L.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));             //라인번호
                    oDS_PS_PP362L.SetValue("U_ColReg01", loopCount, oRecordSet01.Fields.Item("OrdNum").Value);   //작번
                    oDS_PS_PP362L.SetValue("U_ColReg02", loopCount, oRecordSet01.Fields.Item("CardName").Value); //납품처
                    oDS_PS_PP362L.SetValue("U_ColReg03", loopCount, oRecordSet01.Fields.Item("ItemCode").Value); //품목
                    oDS_PS_PP362L.SetValue("U_ColReg04", loopCount, oRecordSet01.Fields.Item("ItemName").Value); //품목명
                    oDS_PS_PP362L.SetValue("U_ColReg05", loopCount, oRecordSet01.Fields.Item("Spec").Value);     //규격
                    oDS_PS_PP362L.SetValue("U_ColReg06", loopCount, oRecordSet01.Fields.Item("Unit").Value);     //단위
                    oDS_PS_PP362L.SetValue("U_ColSum01", loopCount, oRecordSet01.Fields.Item("Quantity").Value); //수량
                    oDS_PS_PP362L.SetValue("U_ColSum02", loopCount, oRecordSet01.Fields.Item("MatAmt").Value);   //자재비
                    oDS_PS_PP362L.SetValue("U_ColSum03", loopCount, oRecordSet01.Fields.Item("DrawAmt").Value);  //설계비
                    oDS_PS_PP362L.SetValue("U_ColSum04", loopCount, oRecordSet01.Fields.Item("GagongAmt").Value);//자체가공비
                    oDS_PS_PP362L.SetValue("U_ColSum05", loopCount, oRecordSet01.Fields.Item("OutgAmt").Value);  //외주가공비
                    oDS_PS_PP362L.SetValue("U_ColSum06", loopCount, oRecordSet01.Fields.Item("OutmAmt").Value);  //외주제작비
                    oDS_PS_PP362L.SetValue("U_ColSum07", loopCount, oRecordSet01.Fields.Item("Total").Value);    //계
                    oDS_PS_PP362L.SetValue("U_ColSum08", loopCount, oRecordSet01.Fields.Item("SjAmt").Value);    //수주금액
                    oDS_PS_PP362L.SetValue("U_ColReg15", loopCount, oRecordSet01.Fields.Item("DocDate").Value);  //수주일자
                    oDS_PS_PP362L.SetValue("U_ColReg16", loopCount, oRecordSet01.Fields.Item("ShipDate").Value); //납기일자
                    oDS_PS_PP362L.SetValue("U_ColReg17", loopCount, oRecordSet01.Fields.Item("EndDate").Value);  //완료일자
                    oDS_PS_PP362L.SetValue("U_ColReg18", loopCount, oRecordSet01.Fields.Item("YesNo").Value);    //완료구분

                    oRecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
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
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
        /// Function ID : PS_PP362_MTX02()
        /// 해당모듈 : PS_PP362
        /// 기능 : Sub작번 TAB 내용 조회
        /// 인수 : prmOrdNum(작번)
        /// 반환값 : 없음
        /// 특이사항 : 없음
        /// </summary>
        /// <param name="prmOrdNum"></param>
        private void PS_PP362_MTX02(string prmOrdNum)
        {
            int loopCount;
            string Query01;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                Query01 = "EXEC PS_PP362_02 '" + prmOrdNum + "'";
                oRecordSet01.DoQuery(Query01);

                oMat02.Clear();
                oMat02.FlushToDataSource();
                oMat02.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    oMat01.Clear();
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }
                for (loopCount = 0; loopCount <= oRecordSet01.RecordCount - 1; loopCount++)
                {
                    if (loopCount != 0)
                    {
                        oDS_PS_PP362M.InsertRecord(loopCount);
                    }
                    oDS_PS_PP362M.Offset = loopCount;

                    oDS_PS_PP362M.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));             //라인번호
                    oDS_PS_PP362M.SetValue("U_ColReg01", loopCount, oRecordSet01.Fields.Item("OrdNum").Value);   //작번
                    oDS_PS_PP362M.SetValue("U_ColReg02", loopCount, oRecordSet01.Fields.Item("OrdSub").Value);   //Sub작번
                    oDS_PS_PP362M.SetValue("U_ColReg03", loopCount, oRecordSet01.Fields.Item("ItemCode").Value); //품목
                    oDS_PS_PP362M.SetValue("U_ColReg04", loopCount, oRecordSet01.Fields.Item("ItemName").Value); //품목명
                    oDS_PS_PP362M.SetValue("U_ColReg05", loopCount, oRecordSet01.Fields.Item("Spec").Value);     //규격
                    oDS_PS_PP362M.SetValue("U_ColReg06", loopCount, oRecordSet01.Fields.Item("Unit").Value);     //단위
                    oDS_PS_PP362M.SetValue("U_ColSum01", loopCount, oRecordSet01.Fields.Item("Quantity").Value); //수량
                    oDS_PS_PP362M.SetValue("U_ColSum02", loopCount, oRecordSet01.Fields.Item("MatAmt").Value);   //자재비
                    oDS_PS_PP362M.SetValue("U_ColSum03", loopCount, oRecordSet01.Fields.Item("DrawAmt").Value);  //설계비
                    oDS_PS_PP362M.SetValue("U_ColSum04", loopCount, oRecordSet01.Fields.Item("GagongAmt").Value);//자체가공비
                    oDS_PS_PP362M.SetValue("U_ColSum05", loopCount, oRecordSet01.Fields.Item("OutgAmt").Value);  //외주가공비
                    oDS_PS_PP362M.SetValue("U_ColSum06", loopCount, oRecordSet01.Fields.Item("OutmAmt").Value);  //외주제작비
                    oDS_PS_PP362M.SetValue("U_ColSum07", loopCount, oRecordSet01.Fields.Item("Total").Value);    //계
                    oDS_PS_PP362M.SetValue("U_ColSum08", loopCount, oRecordSet01.Fields.Item("SjAmt").Value);    //수주금액
                    oDS_PS_PP362M.SetValue("U_ColReg15", loopCount, oRecordSet01.Fields.Item("PP030HNo").Value); //작지등록No

                    oRecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat02.LoadFromDataSource();
                oMat02.AutoResizeColumns();
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
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
            }
        }



        /// <summary>
        /// Function ID : PS_PP362_MTX03()
        /// 해당모듈 : PS_PP362
        /// 기능 : 자재비내역 TAB 내용 조회
        /// 인수 : prmOrdNum(작번), prmOrdSub(Sub작번)
        /// 반환값 : 없음
        /// 특이사항 : 없음
        /// </summary>
        /// <param name="prmOrdNum"></param>
        /// <param name="prmOrdSub"></param>
        private void PS_PP362_MTX03(string prmOrdNum, string prmOrdSub)
        {
            int loopCount;
            string Query01;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                Query01 = "EXEC PS_PP362_03 '" + prmOrdNum + "','" + prmOrdSub + "'";
                oRecordSet01.DoQuery(Query01);

                ProgressBar01.Text = "조회시작!";
                oMat03.Clear();
                oMat03.FlushToDataSource();
                oMat03.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    oMat01.Clear();
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }

                for (loopCount = 0; loopCount <= oRecordSet01.RecordCount - 1; loopCount++)
                {
                    if (loopCount != 0)
                    {
                        oDS_PS_PP362N.InsertRecord(loopCount);
                    }
                    oDS_PS_PP362N.Offset = loopCount;
                    oDS_PS_PP362N.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));              //라인번호
                    oDS_PS_PP362N.SetValue("U_ColReg01", loopCount, oRecordSet01.Fields.Item("Purchase").Value);  //구분
                    oDS_PS_PP362N.SetValue("U_ColReg02", loopCount, oRecordSet01.Fields.Item("PurchaseNm").Value);//구분명
                    oDS_PS_PP362N.SetValue("U_ColReg03", loopCount, oRecordSet01.Fields.Item("OrdNum").Value);    //작번
                    oDS_PS_PP362N.SetValue("U_ColReg04", loopCount, oRecordSet01.Fields.Item("ItemCode").Value);  //품목코드
                    oDS_PS_PP362N.SetValue("U_ColReg05", loopCount, oRecordSet01.Fields.Item("ItemName").Value);  //품명
                    oDS_PS_PP362N.SetValue("U_ColSum01", loopCount, oRecordSet01.Fields.Item("MatAmt").Value);    //자재비

                    oRecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat03.LoadFromDataSource();
                oMat03.AutoResizeColumns();

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
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
            }
        }

        /// <summary>
        /// PS_PP362_MTX04
        /// Function ID : PS_PP362_MTX04()
        /// 해당모듈 : PS_PP362
        /// 기능 : 설계비내역 TAB 내용 조회
        /// 인수 : prmOrdNum(작번), prmOrdSub(Sub작번)
        /// 반환값 : 없음
        /// 특이사항 : 없음
        /// </summary>
        /// <param name="prmOrdNum"></param>
        /// <param name="prmOrdSub"></param>
        private void PS_PP362_MTX04(string prmOrdNum, string prmOrdSub)
        {
            int loopCount;
            string Query01;
            string errMessage = string.Empty;

            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                Query01 = "EXEC PS_PP362_04 '" + prmOrdNum + "','" + prmOrdSub + "'";
                oRecordSet01.DoQuery(Query01);

                ProgressBar01.Text = "조회시작!";
                oMat04.Clear();
                oMat04.FlushToDataSource();
                oMat04.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    oMat01.Clear();
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }
                for (loopCount = 0; loopCount <= oRecordSet01.RecordCount - 1; loopCount++)
                {
                    if (loopCount != 0)
                    {
                        oDS_PS_PP362O.InsertRecord(loopCount);
                    }
                    oDS_PS_PP362O.Offset = loopCount;
                    oDS_PS_PP362O.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));             //라인번호
                    oDS_PS_PP362O.SetValue("U_ColReg01", loopCount, oRecordSet01.Fields.Item("DocDate").Value);  //일자
                    oDS_PS_PP362O.SetValue("U_ColReg02", loopCount, oRecordSet01.Fields.Item("WorkCode").Value); //작업자코드
                    oDS_PS_PP362O.SetValue("U_ColReg03", loopCount, oRecordSet01.Fields.Item("WorkName").Value); //작업자명
                    oDS_PS_PP362O.SetValue("U_ColReg04", loopCount, oRecordSet01.Fields.Item("CpCode").Value);   //공정코드
                    oDS_PS_PP362O.SetValue("U_ColReg05", loopCount, oRecordSet01.Fields.Item("CpName").Value);   //공정명
                    oDS_PS_PP362O.SetValue("U_ColReg06", loopCount, oRecordSet01.Fields.Item("PQty").Value);     //도면매수
                    oDS_PS_PP362O.SetValue("U_ColSum01", loopCount, oRecordSet01.Fields.Item("Amt").Value);      //설계비

                    oRecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat04.LoadFromDataSource();
                oMat04.AutoResizeColumns();
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
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
            }
        }

        /// <summary>
        /// PS_PP362_MTX05
        /// </summary>
        /// <param name="prmOrdNum"></param>
        /// <param name="prmOrdSub"></param>
        private void PS_PP362_MTX05(string prmOrdNum, string prmOrdSub)
        {
            int loopCount;
            string Query01;
            string errMessage = string.Empty;
            
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                Query01 = "EXEC PS_PP362_05 '" + prmOrdNum + "','" + prmOrdSub + "'";
                oRecordSet01.DoQuery(Query01);
                ProgressBar01.Text = "조회시작!";

                oMat05.Clear();
                oMat05.FlushToDataSource();
                oMat05.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    oMat01.Clear();
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }
                for (loopCount = 0; loopCount <= oRecordSet01.RecordCount - 1; loopCount++)
                {
                    if (loopCount != 0)
                    {
                        oDS_PS_PP362P.InsertRecord(loopCount);
                    }
                    oDS_PS_PP362P.Offset = loopCount;
                    oDS_PS_PP362P.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));              //라인번호
                    oDS_PS_PP362P.SetValue("U_ColReg01", loopCount, oRecordSet01.Fields.Item("CpName").Value);    //공정명
                    oDS_PS_PP362P.SetValue("U_ColReg02", loopCount, oRecordSet01.Fields.Item("StdTime").Value);   //표준공수
                    oDS_PS_PP362P.SetValue("U_ColReg03", loopCount, oRecordSet01.Fields.Item("WkTime").Value);    //실동공수
                    oDS_PS_PP362P.SetValue("U_ColReg04", loopCount, oRecordSet01.Fields.Item("ReTime").Value);    //수정공수
                    oDS_PS_PP362P.SetValue("U_ColSum01", loopCount, oRecordSet01.Fields.Item("Amt").Value);       //가공비(실동)
                    oDS_PS_PP362P.SetValue("U_ColReg06", loopCount, oRecordSet01.Fields.Item("Class").Value);     //구분
                    oDS_PS_PP362P.SetValue("U_ColReg07", loopCount, oRecordSet01.Fields.Item("CompltDt").Value);  //완료요구일
                    oDS_PS_PP362P.SetValue("U_ColReg08", loopCount, oRecordSet01.Fields.Item("FirstWkDt").Value); //최초작업
                    oDS_PS_PP362P.SetValue("U_ColReg09", loopCount, oRecordSet01.Fields.Item("LastWkDt").Value);  //최종작업
                    oDS_PS_PP362P.SetValue("U_ColReg10", loopCount, oRecordSet01.Fields.Item("CpCode").Value);    //공정코드
                    oDS_PS_PP362P.SetValue("U_ColReg11", loopCount, oRecordSet01.Fields.Item("OrdNum").Value);    //작번
                    oDS_PS_PP362P.SetValue("U_ColReg12", loopCount, oRecordSet01.Fields.Item("Sub1_2").Value);    //서브작번

                    oRecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat05.LoadFromDataSource();
                oMat05.AutoResizeColumns();
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
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
            }
        }

        /// <summary>
        /// PS_PP362_MTX06
        /// Function ID : PS_PP362_MTX06()
        /// 해당모듈 : PS_PP362
        /// 기능 : 자체가공비내역(작업자별내역) TAB 내용 조회
        /// 인수 : prmOrdNum(작번), prmOrdSub(Sub작번), prmCpCode(공정코드)
        /// 반환값 : 없음
        /// 특이사항 : 없음
        /// </summary>
        /// <param name="prmOrdNum"></param>
        /// <param name="prmOrdSub"></param>
        /// <param name="prmCpCode"></param>
        private void PS_PP362_MTX06(string prmOrdNum, string prmOrdSub, string prmCpCode)
        {
            int loopCount;
            string Query01;
            string errMessage = string.Empty;
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                Query01 = "EXEC PS_PP362_06 '" + prmOrdNum + "','" + prmOrdSub + "','" + prmCpCode + "'";
                oRecordSet01.DoQuery(Query01);
                ProgressBar01.Text = "조회시작!";

                oMat06.Clear();
                oMat06.FlushToDataSource();
                oMat06.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    oMat01.Clear();
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }
                for (loopCount = 0; loopCount <= oRecordSet01.RecordCount - 1; loopCount++)
                {
                    if (loopCount != 0)
                    {
                        oDS_PS_PP362Q.InsertRecord(loopCount);
                    }
                    oDS_PS_PP362Q.Offset = loopCount;
                    oDS_PS_PP362Q.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));            //라인번호
                    oDS_PS_PP362Q.SetValue("U_ColReg01", loopCount, oRecordSet01.Fields.Item("WorkDt").Value);  //작업일자
                    oDS_PS_PP362Q.SetValue("U_ColReg02", loopCount, oRecordSet01.Fields.Item("EmpCode").Value); //사번
                    oDS_PS_PP362Q.SetValue("U_ColReg03", loopCount, oRecordSet01.Fields.Item("EmpName").Value); //성명
                    oDS_PS_PP362Q.SetValue("U_ColQty01", loopCount, oRecordSet01.Fields.Item("WkTime").Value);  //공수
                    oDS_PS_PP362Q.SetValue("U_ColReg05", loopCount, oRecordSet01.Fields.Item("Qty").Value);     //완료수량
                    oDS_PS_PP362Q.SetValue("U_ColReg06", loopCount, oRecordSet01.Fields.Item("Price").Value);   //단가
                    oDS_PS_PP362Q.SetValue("U_ColSum02", loopCount, oRecordSet01.Fields.Item("Amt").Value);     //금액(가공비)
                    oDS_PS_PP362Q.SetValue("U_ColReg08", loopCount, oRecordSet01.Fields.Item("Class").Value);   //구분
                    oDS_PS_PP362Q.SetValue("U_ColReg09", loopCount, oRecordSet01.Fields.Item("CpCode").Value);  //공정코드
                    oDS_PS_PP362Q.SetValue("U_ColReg10", loopCount, oRecordSet01.Fields.Item("CpName").Value);  //공정명
                    oDS_PS_PP362Q.SetValue("U_ColReg11", loopCount, oRecordSet01.Fields.Item("PP040HNo").Value);//작업일보번호

                    oRecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat06.LoadFromDataSource();
                oMat06.AutoResizeColumns();
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
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
            }
        }

        /// <summary>
        /// PS_PP362_MTX07
        /// Function ID : PS_PP362_MTX07()
        /// 해당모듈 : PS_PP362
        /// 기능 : 외주가공비내역 TAB 내용 조회
        /// 인수 : prmOrdNum(작번), prmOrdSub(Sub작번)
        /// 반환값 : 없음
        /// 특이사항 : 없음
        /// </summary>
        /// <param name="prmOrdNum"></param>
        /// <param name="prmOrdSub"></param>
        private void PS_PP362_MTX07(string prmOrdNum, string prmOrdSub)
        {
            int loopCount;
            string Query01;
            string errMessage = string.Empty;

            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                Query01 = "EXEC PS_PP362_07 '" + prmOrdNum + "','" + prmOrdSub + "'";
                oRecordSet01.DoQuery(Query01);
                ProgressBar01.Text = "조회시작!";

                oMat07.Clear();
                oMat07.FlushToDataSource();
                oMat07.LoadFromDataSource();
                if (oRecordSet01.RecordCount == 0)
                {
                    oMat01.Clear();
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }
                for (loopCount = 0; loopCount <= oRecordSet01.RecordCount - 1; loopCount++)
                {
                    if (loopCount != 0)
                    {
                        oDS_PS_PP362R.InsertRecord(loopCount);
                    }
                    oDS_PS_PP362R.Offset = loopCount;
                    oDS_PS_PP362R.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));            //라인번호
                    oDS_PS_PP362R.SetValue("U_ColReg01", loopCount, oRecordSet01.Fields.Item("CpCode").Value);  //공정코드
                    oDS_PS_PP362R.SetValue("U_ColReg02", loopCount, oRecordSet01.Fields.Item("CpName").Value);  //공정명
                    oDS_PS_PP362R.SetValue("U_ColReg03", loopCount, oRecordSet01.Fields.Item("OrdNum").Value);  //작번
                    oDS_PS_PP362R.SetValue("U_ColReg04", loopCount, oRecordSet01.Fields.Item("ItemName").Value);//품명
                    oDS_PS_PP362R.SetValue("U_ColSum01", loopCount, oRecordSet01.Fields.Item("Amt").Value);     //금액

                    oRecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat07.LoadFromDataSource();
                oMat07.AutoResizeColumns();
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
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
            }
        }


        /// <summary>
        /// PS_PP362_MTX08
        /// Function ID : PS_PP362_MTX08()
        /// 해당모듈 : PS_PP362
        /// 기능 : 외주제작비내역 TAB 내용 조회
        /// 인수 : prmOrdNum(작번), prmOrdSub(Sub작번)
        /// 반환값 : 없음
        /// 특이사항 : 없음
        /// </summary>
        /// <param name="prmOrdNum"></param>
        /// <param name="prmOrdSub"></param>
        private void PS_PP362_MTX08(string prmOrdNum, string prmOrdSub)
        {
            int loopCount = 0;
            string Query01 = null;
            string errMessage = string.Empty;

            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                Query01 = "EXEC PS_PP362_08 '" + prmOrdNum + "','" + prmOrdSub + "'";
                oRecordSet01.DoQuery(Query01);
                ProgressBar01.Text = "조회시작!";

                oMat08.Clear();
                oMat08.FlushToDataSource();
                oMat08.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    oMat01.Clear();
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }
                for (loopCount = 0; loopCount <= oRecordSet01.RecordCount - 1; loopCount++)
                {
                    if (loopCount != 0)
                    {
                        oDS_PS_PP362S.InsertRecord(loopCount);
                    }
                    oDS_PS_PP362S.Offset = loopCount;

                    oDS_PS_PP362S.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));            //라인번호
                    oDS_PS_PP362S.SetValue("U_ColReg01", loopCount, oRecordSet01.Fields.Item("CpCode").Value);  //공정코드
                    oDS_PS_PP362S.SetValue("U_ColReg02", loopCount, oRecordSet01.Fields.Item("CpName").Value);  //공정명
                    oDS_PS_PP362S.SetValue("U_ColReg03", loopCount, oRecordSet01.Fields.Item("OrdNum").Value);  //작번
                    oDS_PS_PP362S.SetValue("U_ColReg04", loopCount, oRecordSet01.Fields.Item("ItemName").Value);//품명
                    oDS_PS_PP362S.SetValue("U_ColSum01", loopCount, oRecordSet01.Fields.Item("Amt").Value);     //금액

                    oRecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat08.LoadFromDataSource();
                oMat08.AutoResizeColumns();
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
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_PP362_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP362'", "");
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
        /// 리포트 조회
        /// </summary>
        [STAThread]
        private void PS_PP362_Print_Report(object prmItemUID)
        {
            string WinTitle;
            string ReportName;
            string FrDt;     //작번등록년월(시작)
            string ToDt;     //작번등록년월(종료)
            string FrgnName; //품명
            string CardType; //거래처구분
            string SPEC;     //규격
            string ItemClass;//품목구분
            string WCYN;     //생산완료여부
            string CardCode; //거래처
            string ItemCode; //품목(작번)
            string OrdNum;   //Main작번
            string OrdSub;   //Sub작번
            string CpCode;   //공정코드
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                WinTitle = "[PS_PP362] 레포트";
                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List

                if (prmItemUID.ToString() == "Btn01")
                {
                    FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
                    ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();
                    FrgnName = (string.IsNullOrEmpty(oForm.Items.Item("FrgnName").Specific.Value) ? "%" : oForm.Items.Item("FrgnName").Specific.Value);
                    CardType = oForm.Items.Item("CardType").Specific.Selected.Value.ToString().Trim();
                    SPEC = (string.IsNullOrEmpty(oForm.Items.Item("Spec").Specific.Value) ? "%" : oForm.Items.Item("Spec").Specific.Value);
                    ItemClass = oForm.Items.Item("ItemClass").Specific.Selected.Value.ToString().Trim();
                    WCYN = (string.IsNullOrEmpty(oForm.Items.Item("WCYN").Specific.Selected.Value) ? "%" : oForm.Items.Item("WCYN").Specific.Selected.Value);
                    CardCode = (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value) ? "%" : oForm.Items.Item("CardCode").Specific.Value);
                    ItemCode = (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value) ? "%" : oForm.Items.Item("ItemCode").Specific.Value);

                    ReportName = "PS_PP362_01.rpt";

                    dataPackParameter.Add(new PSH_DataPackClass("@FrDt", FrDt));
                    dataPackParameter.Add(new PSH_DataPackClass("@ToDt", ToDt));
                    dataPackParameter.Add(new PSH_DataPackClass("@FrgnName", FrgnName));
                    dataPackParameter.Add(new PSH_DataPackClass("@CardType", CardType));
                    dataPackParameter.Add(new PSH_DataPackClass("@Spec", SPEC));
                    dataPackParameter.Add(new PSH_DataPackClass("@ItemClass", ItemClass));
                    dataPackParameter.Add(new PSH_DataPackClass("@WCYN", WCYN));
                    dataPackParameter.Add(new PSH_DataPackClass("@CardCode", CardCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@ItemCode", ItemCode));

                    formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
                }
                else if (prmItemUID.ToString() == "Btn02")
                {
                    OrdNum = oMat02.Columns.Item("OrdNum").Cells.Item(1).Specific.Value;

                    ReportName = "PS_PP362_02.rpt";
                    dataPackParameter.Add(new PSH_DataPackClass("@OrdNum", OrdNum));

                    formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
                }
                else if (prmItemUID.ToString() == "Btn03")
                {
                    OrdNum = oForm.Items.Item("MainJak").Specific.Value.ToString().Trim();
                    OrdSub = oForm.Items.Item("SubJak").Specific.Value.ToString().Trim();

                    ReportName = "PS_PP362_03.rpt";
                    dataPackParameter.Add(new PSH_DataPackClass("@OrdNum", OrdNum));
                    dataPackParameter.Add(new PSH_DataPackClass("@Sub1_2", OrdSub));

                    formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
                }
                else if (prmItemUID.ToString() == "Btn04")
                {
                    OrdNum = oForm.Items.Item("MainJak").Specific.Value.ToString().Trim();
                    OrdSub = oForm.Items.Item("SubJak").Specific.Value.ToString().Trim();

                    ReportName = "PS_PP362_04.rpt";
                    dataPackParameter.Add(new PSH_DataPackClass("@OrdNum", OrdNum));
                    dataPackParameter.Add(new PSH_DataPackClass("@Sub1_2", OrdSub));

                    formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
                }
                else if (prmItemUID.ToString() == "Btn05")
                {
                    OrdNum = oForm.Items.Item("MainJak").Specific.Value.ToString().Trim();
                    OrdSub = oForm.Items.Item("SubJak").Specific.Value.ToString().Trim();

                    ReportName = "PS_PP362_05.rpt";
                    dataPackParameter.Add(new PSH_DataPackClass("@OrdNum", OrdNum));
                    dataPackParameter.Add(new PSH_DataPackClass("@Sub1_2", OrdSub));

                    formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
                }
                else if (prmItemUID.ToString() == "Btn06")
                {
                    OrdNum = oForm.Items.Item("MainJak").Specific.Value.ToString().Trim();
                    OrdSub = oForm.Items.Item("SubJak").Specific.Value.ToString().Trim();
                    CpCode = oForm.Items.Item("CpCode").Specific.Value.ToString().Trim();

                    ReportName = "PS_PP362_06.rpt";
                    dataPackParameter.Add(new PSH_DataPackClass("@OrdNum", OrdNum));
                    dataPackParameter.Add(new PSH_DataPackClass("@Sub1_2", OrdSub));
                    dataPackParameter.Add(new PSH_DataPackClass("@CpCode", CpCode));

                    formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
                }
                else if (prmItemUID.ToString() == "Btn07")
                {
                    OrdNum = oForm.Items.Item("MainJak").Specific.Value.ToString().Trim();
                    OrdSub = oForm.Items.Item("SubJak").Specific.Value.ToString().Trim();

                    ReportName = "PS_PP362_07.rpt";
                    dataPackParameter.Add(new PSH_DataPackClass("@OrdNum", OrdNum));
                    dataPackParameter.Add(new PSH_DataPackClass("@Sub1_2", OrdSub));

                    formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
                }
                else if (prmItemUID.ToString() == "Btn08")
                {
                    OrdNum = oForm.Items.Item("MainJak").Specific.Value.ToString().Trim();
                    OrdSub = oForm.Items.Item("SubJak").Specific.Value.ToString().Trim();

                    ReportName = "PS_PP362_08.rpt";
                    dataPackParameter.Add(new PSH_DataPackClass("@OrdNum", OrdNum));
                    dataPackParameter.Add(new PSH_DataPackClass("@Sub1_2", OrdSub));

                    formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
                }
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

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                //    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "btnSearch")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            oForm.Items.Item("MainJak").Specific.Value = "";
                            oForm.Items.Item("SubJak").Specific.Value = "";

                            oMat01.Clear();
                            oDS_PS_PP362L.Clear();

                            //다른 TAB의 정보 초기화
                            oMat02.Clear();
                            oDS_PS_PP362M.Clear();
                            oMat03.Clear();
                            oDS_PS_PP362N.Clear();
                            oMat04.Clear();
                            oDS_PS_PP362O.Clear();
                            oMat05.Clear();
                            oDS_PS_PP362P.Clear();
                            oMat06.Clear();
                            oDS_PS_PP362Q.Clear();
                            oMat07.Clear();
                            oDS_PS_PP362R.Clear();
                            oMat08.Clear();
                            oDS_PS_PP362S.Clear();

                            oForm.Items.Item("Folder01").Specific.Select(); //Main작번 TAB 선택
                            PS_PP362_MTX01();  //매트릭스에 데이터 로드
                        }
                    }
                    else if (pVal.ItemUID == "Btn01" || pVal.ItemUID == "Btn02" || pVal.ItemUID == "Btn03" || pVal.ItemUID == "Btn04" || pVal.ItemUID == "Btn05" || pVal.ItemUID == "Btn06" || pVal.ItemUID == "Btn07" || pVal.ItemUID == "Btn08")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            System.Threading.Thread thread = new System.Threading.Thread(new System.Threading.ParameterizedThreadStart(PS_PP362_Print_Report));
                            thread.SetApartmentState(System.Threading.ApartmentState.STA);
                            thread.Start(pVal.ItemUID);
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    
                    if (pVal.ItemUID == "Folder01")//Folder01이 선택되었을 때
                    {
                        oForm.PaneLevel = 1;
                    }
                    if (pVal.ItemUID == "Folder02")//Folder02가 선택되었을 때
                    {
                        oForm.PaneLevel = 2;
                    }
                    if (pVal.ItemUID == "Folder03")//Folder03이 선택되었을 때
                    {
                        oForm.PaneLevel = 3;
                    }
                    if (pVal.ItemUID == "Folder04") //Folder04이 선택되었을 때
                    {
                        oForm.PaneLevel = 4;
                    }
                    if (pVal.ItemUID == "Folder05") //Folder05이 선택되었을 때
                    {
                        oForm.PaneLevel = 5;
                    }
                    if (pVal.ItemUID == "Folder06") //Folder06이 선택되었을 때
                    {
                        oForm.PaneLevel = 6;
                    }
                    if (pVal.ItemUID == "Folder07") //Folder07이 선택되었을 때
                    {
                        oForm.PaneLevel = 7;
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", ""); //거래처코드 포맷서치 활성
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", ""); //품목코드(작번) 포맷서치 활성
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
                else if (pVal.ItemUID == "Mat05")
                {
                    if (pVal.Row > 0)
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = pVal.ColUID;
                        oLastColRow01 = pVal.Row;
                    }
                }
                else if (pVal.ItemUID == "Mat06")
                {
                    if (pVal.Row > 0)
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = pVal.ColUID;
                        oLastColRow01 = pVal.Row;
                    }
                }
                else if (pVal.ItemUID == "Mat07")
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
                    else if (pVal.ItemUID == "Mat05")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat05.SelectRow(pVal.Row, true, false);
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                    }
                    else if (pVal.ItemUID == "Mat06")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat06.SelectRow(pVal.Row, true, false);
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                    }
                    else if (pVal.ItemUID == "Mat07")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat07.SelectRow(pVal.Row, true, false);
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                    }
                    else if (pVal.ItemUID == "Mat08")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat08.SelectRow(pVal.Row, true, false);
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
            string oQuery01;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "CardCode")
                        {
                            oQuery01 = "SELECT CardName, CardCode FROM [OCRD] WHERE CardCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'";
                            oRecordSet01.DoQuery(oQuery01);
                            oForm.Items.Item("CardName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }
                        else if (pVal.ItemUID == "ItemCode")
                        {
                            oQuery01 = "SELECT FrgnName, ItemCode FROM [OITM] WHERE ItemCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'";
                            oRecordSet01.DoQuery(oQuery01);
                            oForm.Items.Item("ItemName").Specific.Value =oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }
                        oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                BubbleEvent = false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
                oForm.Freeze(false);
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
                        PS_PP030 PS_PP030 = new PS_PP030();
                        PS_PP030.LoadForm(oMat01.Columns.Item("PP030HNo").Cells.Item(pVal.Row).Specific.String);
                        BubbleEvent = false;
                    }
                    else if (pVal.ItemUID == "Mat06")
                    {
                        PS_PP040 PS_PP040 = new PS_PP040();
                        PS_PP040.LoadForm(oMat06.Columns.Item("PP040HNo").Cells.Item(pVal.Row).Specific.Value);
                        BubbleEvent = false;
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat03);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat04);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat05);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat06);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat07);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat08);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP362L);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP362M);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP362N);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP362O);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP362P);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP362Q);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP362R);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP362S);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Raise_EVENT_DOUBLE_CLICK 이벤트
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
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row == 0)
                        {
                            //정렬
                            oMat01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
                            oMat01.FlushToDataSource();
                        }
                        else
                        {
                            //다른 TAB의 정보 초기화
                            oMat02.Clear();
                            oDS_PS_PP362M.Clear();
                            oMat03.Clear();
                            oDS_PS_PP362N.Clear();
                            oMat04.Clear();
                            oDS_PS_PP362O.Clear();
                            oMat05.Clear();
                            oDS_PS_PP362P.Clear();
                            oMat06.Clear();
                            oDS_PS_PP362Q.Clear();
                            oMat07.Clear();
                            oDS_PS_PP362R.Clear();
                            oMat08.Clear();
                            oDS_PS_PP362S.Clear();

                            oForm.Items.Item("Folder02").Specific.Select();  //Sub작번 TAB 선택
                            PS_PP362_MTX02(oMat01.Columns.Item("OrdNum").Cells.Item(pVal.Row).Specific.Value);
                        }
                    }
                    else if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.Row == 0)
                        {
                            //정렬
                            oMat02.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
                            oMat02.FlushToDataSource();
                        }
                        else
                        {
                            oForm.Items.Item("MainJak").Specific.Value = oMat02.Columns.Item("OrdNum").Cells.Item(pVal.Row).Specific.Value;
                            oForm.Items.Item("SubJak").Specific.Value = oMat02.Columns.Item("OrdSub").Cells.Item(pVal.Row).Specific.Value;

                            //다른 TAB의 정보 초기화
                            oMat03.Clear();
                            oDS_PS_PP362N.Clear();
                            oMat04.Clear();
                            oDS_PS_PP362O.Clear();
                            oMat05.Clear();
                            oDS_PS_PP362P.Clear();
                            oMat06.Clear();
                            oDS_PS_PP362Q.Clear();
                            oMat07.Clear();
                            oDS_PS_PP362R.Clear();
                            oMat08.Clear();
                            oDS_PS_PP362S.Clear();
                            
                            PS_PP362_MTX03(oMat02.Columns.Item("OrdNum").Cells.Item(pVal.Row).Specific.Value, oMat02.Columns.Item("OrdSub").Cells.Item(pVal.Row).Specific.Value);//자재비
                            PS_PP362_MTX04(oMat02.Columns.Item("OrdNum").Cells.Item(pVal.Row).Specific.Value, oMat02.Columns.Item("OrdSub").Cells.Item(pVal.Row).Specific.Value);//설계비
                            PS_PP362_MTX05(oMat02.Columns.Item("OrdNum").Cells.Item(pVal.Row).Specific.Value, oMat02.Columns.Item("OrdSub").Cells.Item(pVal.Row).Specific.Value);//자체가공비_공정별
                            //자체가공비_작업자별
                            PS_PP362_MTX07(oMat02.Columns.Item("OrdNum").Cells.Item(pVal.Row).Specific.Value, oMat02.Columns.Item("OrdSub").Cells.Item(pVal.Row).Specific.Value);//외주가공비
                            PS_PP362_MTX08(oMat02.Columns.Item("OrdNum").Cells.Item(pVal.Row).Specific.Value, oMat02.Columns.Item("OrdSub").Cells.Item(pVal.Row).Specific.Value);//외주제작비
                        }
                    }
                    else if (pVal.ItemUID == "Mat03")
                    {

                        if (pVal.Row == 0)
                        {
                            oMat03.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
                            oMat03.FlushToDataSource();
                        }
                    }
                    else if (pVal.ItemUID == "Mat04")
                    {
                        if (pVal.Row == 0)
                        {
                            oMat04.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
                            oMat04.FlushToDataSource();
                        }
                    }
                    else if (pVal.ItemUID == "Mat05")
                    {
                        if (pVal.Row == 0)
                        {
                            oMat05.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
                            oMat05.FlushToDataSource();
                        }
                        else
                        {
                            oForm.Items.Item("CpCode").Specific.Value = oMat05.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value;
                            PS_PP362_MTX06(oMat05.Columns.Item("OrdNum").Cells.Item(pVal.Row).Specific.Value, oMat05.Columns.Item("Sub1_2").Cells.Item(pVal.Row).Specific.Value, oMat05.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value);
                        }
                    }
                    else if (pVal.ItemUID == "Mat06")
                    {
                        if (pVal.Row == 0)
                        {
                            oMat06.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
                            oMat06.FlushToDataSource();
                        }
                    }
                    else if (pVal.ItemUID == "Mat07")
                    {
                        if (pVal.Row == 0)
                        {
                            oMat07.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
                            oMat07.FlushToDataSource();
                        }
                    }
                    else if (pVal.ItemUID == "Mat08")
                    {
                        if (pVal.Row == 0)
                        {
                            oMat08.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
                            oMat08.FlushToDataSource();
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
        /// RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PS_PP362_FormResize();
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
                        case "7169":
                            //엑셀 내보내기

                            //엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                            PS_PP362_AddMatrixRow(oMat01.VisualRowCount, oMat01, oDS_PS_PP362L, false); //Main작번
                            PS_PP362_AddMatrixRow(oMat02.VisualRowCount, oMat02, oDS_PS_PP362M, false); //Sub작번
                            PS_PP362_AddMatrixRow(oMat03.VisualRowCount, oMat03, oDS_PS_PP362N, false); //자재비내역
                            PS_PP362_AddMatrixRow(oMat04.VisualRowCount, oMat04, oDS_PS_PP362O, false); //설계비내역
                            PS_PP362_AddMatrixRow(oMat05.VisualRowCount, oMat05, oDS_PS_PP362P, false); //자체가공비내역_공정별
                            PS_PP362_AddMatrixRow(oMat06.VisualRowCount, oMat06, oDS_PS_PP362Q, false); //자체가공비내역_작업자별
                            PS_PP362_AddMatrixRow(oMat07.VisualRowCount, oMat07, oDS_PS_PP362R, false); //외주가공비내역
                            PS_PP362_AddMatrixRow(oMat08.VisualRowCount, oMat08, oDS_PS_PP362S, false); //외주제작비내역
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
                        case "1282": //추가
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
                        case "1287": //복제
                            break;
                        case "7169":
                            //엑셀 내보내기

                            //엑셀 내보내기 이후 처리
                            oForm.Freeze(true);
                            
                            oDS_PS_PP362L.RemoveRecord(oDS_PS_PP362L.Size - 1); //MAIN작번
                            oMat01.LoadFromDataSource();
                            oDS_PS_PP362M.RemoveRecord(oDS_PS_PP362M.Size - 1); //Sub작번
                            oMat02.LoadFromDataSource();
                            oDS_PS_PP362N.RemoveRecord(oDS_PS_PP362N.Size - 1); //자재비내역
                            oMat03.LoadFromDataSource();                        
                            oDS_PS_PP362O.RemoveRecord(oDS_PS_PP362O.Size - 1); //설계비내역
                            oMat04.LoadFromDataSource();                        
                            oDS_PS_PP362P.RemoveRecord(oDS_PS_PP362P.Size - 1); //자체가공비내역_공정별
                            oMat05.LoadFromDataSource();                        
                            oDS_PS_PP362Q.RemoveRecord(oDS_PS_PP362Q.Size - 1); //자체가공비내역_작업자별
                            oMat06.LoadFromDataSource();                        
                            oDS_PS_PP362R.RemoveRecord(oDS_PS_PP362R.Size - 1); //외주가공비내역
                            oMat07.LoadFromDataSource();                        
                            oDS_PS_PP362S.RemoveRecord(oDS_PS_PP362S.Size - 1); //외주제작비내역
                            oMat08.LoadFromDataSource();
                            oForm.Freeze(false);
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
                else if (pVal.ItemUID == "Mat05")
                {
                    if (pVal.Row > 0)
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = pVal.ColUID;
                        oLastColRow01 = pVal.Row;
                    }
                }
                else if (pVal.ItemUID == "Mat06")
                {
                    if (pVal.Row > 0)
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = pVal.ColUID;
                        oLastColRow01 = pVal.Row;
                    }
                }
                else if (pVal.ItemUID == "Mat07")
                {
                    if (pVal.Row > 0)
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = pVal.ColUID;
                        oLastColRow01 = pVal.Row;
                    }
                }
                else if (pVal.ItemUID == "Mat08")
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
    }
}
