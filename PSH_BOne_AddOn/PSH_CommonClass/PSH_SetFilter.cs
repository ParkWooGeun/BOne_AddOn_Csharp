namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 이벤트 필터(2021.03.05부 사용안함, 각 화면 클래스에서 동적으로 이벤트 필터 생성)
	/// </summary>
	internal static class PSH_SetFilter
    {
        public static void Execute()
		{
			SAPbouiCOM.EventFilters oFilters = null;
			SAPbouiCOM.EventFilter oFilter = null;

			oFilters = new SAPbouiCOM.EventFilters();

			oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK); //Main Menu 클릭 이벤트를 실행하기 위한 필수 이벤트(모든 클래스 필터 적용)

			//PSH_Globals.ExecuteEventFilter(typeof(PS_CO001), ref oFilters, ref oFilter);
			//PSH_Globals.ExecuteEventFilter(typeof(PS_SD600), ref oFilters, ref oFilter);
			//PSH_Globals.ExecuteEventFilter(typeof(PS_SD602), ref oFilters, ref oFilter);
			//PSH_Globals.ExecuteEventFilter(typeof(PS_SD603), ref oFilters, ref oFilter);

			//ITEM_PRESSED(ref oFilter, ref oFilters); //1
			//KEY_DOWN(ref oFilter, ref oFilters); //2
			//GOT_FOCUS(ref oFilter, ref oFilters); //3
			//LOST_FOCUS(ref oFilter, ref oFilters); //4
			//COMBO_SELECT(ref oFilter, ref oFilters); //5
			//CLICK(ref oFilter, ref oFilters); //6
			//DOUBLE_CLICK(ref oFilter, ref oFilters); //7
			//MATRIX_LINK_PRESSED(ref oFilter, ref oFilters); //8
			//MATRIX_COLLAPSE_PRESSED(ref oFilter, ref oFilters); //9
			//VALIDATE(ref oFilter, ref oFilters); //10
			//MATRIX_LOAD(ref oFilter, ref oFilters); //11
			//DATASOURCE_LOAD(ref oFilter, ref oFilters); //12
			//FORM_LOAD(ref oFilter, ref oFilters); //16
			//FORM_UNLOAD(ref oFilter, ref oFilters); //17
			//FORM_ACTIVATE(ref oFilter, ref oFilters); //18
			//FORM_DEACTIVATE(ref oFilter, ref oFilters); //19
			//FORM_CLOSE(ref oFilter, ref oFilters); //20
			//FORM_RESIZE(ref oFilter, ref oFilters); //21
			//FORM_KEY_DOWN(ref oFilter, ref oFilters); //22
			//FORM_MENU_HILIGHT(ref oFilter, ref oFilters); //23
			//PRINT(ref oFilter, ref oFilters); //24
			//PRINT_DATA(ref oFilter, ref oFilters); //25
			//CHOOSE_FROM_LIST(ref oFilter, ref oFilters); //27
			//RIGHT_CLICK(ref oFilter, ref oFilters); //28
			//MENU_CLICK(ref oFilter, ref oFilters); //32 (Main Menu 클릭 이벤트를 실행하기 위한 필수 이벤트필터)
			//FORM_DATA_ADD(ref oFilter, ref oFilters); //33
			//FORM_DATA_UPDATE(ref oFilter, ref oFilters); //34
			//FORM_DATA_DELETE(ref oFilter, ref oFilters); //35
			//FORM_DATA_LOAD(ref oFilter, ref oFilters); //36

			//Setting the application with the EventFilters object
			PSH_Globals.SBO_Application.SetFilter(oFilters);
			
			oFilter = null;
			oFilters = null;
		}

		//private static void ITEM_PRESSED(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);
		//	//oFilter.AddEx("");
		//}

		//private static void KEY_DOWN(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN);
		//	//oFilter.AddEx("");
		//}

		//private static void GOT_FOCUS(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS);
		//	//oFilter.AddEx("");
		//}

		//private static void LOST_FOCUS(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS);
		//	//oFilter.AddEx("");
		//}

		//private static void COMBO_SELECT(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT);
		//	//oFilter.AddEx("");
		//}	

		//private static void CLICK(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
		//	//oFilter.AddEx("");
		//}

		//private static void DOUBLE_CLICK(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK);
		//	//oFilter.AddEx("");
		//}

		//private static void MATRIX_LINK_PRESSED(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED);
		//	//oFilter.AddEx("");
		//}

		//private static void MATRIX_COLLAPSE_PRESSED(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
  //      {
  //          oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED);
		//	//oFilter.AddEx("");
		//}

		//private static void VALIDATE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE);
		//	//oFilter.AddEx("");
		//}

		//private static void MATRIX_LOAD(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD);
		//	//oFilter.AddEx("");
		//}

		//private static void DATASOURCE_LOAD(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD);
		//	//oFilter.AddEx("");
		//}

		//private static void FORM_LOAD(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);
		//	//oFilter.AddEx("");
		//}

		//private static void FORM_UNLOAD(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD);
		//	//oFilter.AddEx("");
		//}

		//private static void FORM_ACTIVATE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE);
		//	//oFilter.AddEx("");
		//}

		//private static void FORM_DEACTIVATE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE);
		//	//oFilter.AddEx("");
		//}

		//private static void FORM_CLOSE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE);
		//	//oFilter.AddEx("");
		//}

		//private static void FORM_RESIZE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_RESIZE);
		//	//oFilter.AddEx("");
		//}

		//private static void FORM_KEY_DOWN(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN);
		//	//oFilter.AddEx("");
		//}

		//private static void FORM_MENU_HILIGHT(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT);
		//	//oFilter.AddEx("");
		//}

		//private static void PRINT(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_PRINT);
		//	//oFilter.AddEx("");
		//}

		//private static void PRINT_DATA(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_PRINT_DATA);
		//	//oFilter.AddEx("");
		//}

		//private static void CHOOSE_FROM_LIST(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST);
		//	//oFilter.AddEx("");
		//}

		//private static void RIGHT_CLICK(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK);
		//	//oFilter.AddEx("");
		//}

		//private static void MENU_CLICK(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
		//	//oFilter.AddEx("");
		//}

		//private static void FORM_DATA_ADD(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD);
		//	//oFilter.AddEx("");
		//}

		//private static void FORM_DATA_UPDATE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE);
		//	//oFilter.AddEx("");
		//}

		//private static void FORM_DATA_DELETE(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE);
		//	//oFilter.AddEx("");
		//}

		//private static void FORM_DATA_LOAD(ref SAPbouiCOM.EventFilter oFilter, ref SAPbouiCOM.EventFilters oFilters)
		//{
		//	oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD);
		//	//oFilter.AddEx("");
		//}
	}
}
