#include "MyForm.h"
#include <Windows.h>

using namespace System;
using namespace System::Windows::Forms;


[STAThread]
int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR    lpCmdLine, int       nCmdShow)	// int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR    lpCmdLine, int       nCmdShow) <- for Release		//void Main(array<String^>^ args) <- for Debug//

{
	Application::EnableVisualStyles();
	Application::SetCompatibleTextRenderingDefault(false);

	Data_Base_2::MyForm form;
	Application::Run(%form);


	return(0);
}
