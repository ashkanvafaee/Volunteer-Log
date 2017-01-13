#define _CRT_SECURE_NO_WARNINGS
#include <stdio.h>
#include <stdlib.h>
#include <string>
#include <fstream>
#include <iostream>
#include "DataBox.h"
#include "List.h"
#include <sys/stat.h>
#include "listEdit.h"





#pragma once

namespace Data_Base_2 {


	using namespace NetOffice::ExcelApi;
	using namespace NetOffice::ExcelApi::Enums;
	using namespace NetOffice::ExcelApi::Tools::Utils;
	using namespace System::IO;
	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// Summary for MyForm
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
			//
			//TODO: Add the constructor code here
			//
		}

	protected:
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}




	private: System::Windows::Forms::Button^  button1;
	private: System::Windows::Forms::Button^  button2;



	private: System::Windows::Forms::ToolStripMenuItem^  fileToolStripMenuItem;
	private: System::Windows::Forms::ToolStripMenuItem^  newLogToolStripMenuItem;
	private: System::Windows::Forms::ToolStripMenuItem^  exitToolStripMenuItem;
	private: System::Windows::Forms::MenuStrip^  menuStrip1;

	private: System::Windows::Forms::Label^  label1;

	private: System::Windows::Forms::ToolStripMenuItem^  editToolStripMenuItem;
	private: System::Windows::Forms::ToolStripMenuItem^  memberInfoToolStripMenuItem;
	private: System::Windows::Forms::PictureBox^  pictureBox2;
	private: System::Windows::Forms::PictureBox^  pictureBox1;
	private: System::Windows::Forms::Panel^  panel1;
	private: System::Windows::Forms::Panel^  panel2;
	private: System::Windows::Forms::Panel^  panel3;
	private: System::Windows::Forms::Panel^  panel4;
	private: System::Windows::Forms::Panel^  panel5;
	private: System::Windows::Forms::Panel^  panel6;
	private: System::Windows::Forms::Panel^  panel7;






















	protected:

	private:
		/// <summary>
		/// Required designer variable.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		void InitializeComponent(void)
		{
			System::ComponentModel::ComponentResourceManager^  resources = (gcnew System::ComponentModel::ComponentResourceManager(MyForm::typeid));
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->button2 = (gcnew System::Windows::Forms::Button());
			this->fileToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->newLogToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->exitToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->menuStrip1 = (gcnew System::Windows::Forms::MenuStrip());
			this->editToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->memberInfoToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->pictureBox2 = (gcnew System::Windows::Forms::PictureBox());
			this->pictureBox1 = (gcnew System::Windows::Forms::PictureBox());
			this->panel1 = (gcnew System::Windows::Forms::Panel());
			this->panel5 = (gcnew System::Windows::Forms::Panel());
			this->panel2 = (gcnew System::Windows::Forms::Panel());
			this->panel3 = (gcnew System::Windows::Forms::Panel());
			this->panel4 = (gcnew System::Windows::Forms::Panel());
			this->panel6 = (gcnew System::Windows::Forms::Panel());
			this->panel7 = (gcnew System::Windows::Forms::Panel());
			this->menuStrip1->SuspendLayout();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox2))->BeginInit();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox1))->BeginInit();
			this->panel1->SuspendLayout();
			this->panel6->SuspendLayout();
			this->SuspendLayout();
			// 
			// button1
			// 
			this->button1->AutoSize = true;
			this->button1->BackColor = System::Drawing::Color::Transparent;
			this->button1->BackgroundImageLayout = System::Windows::Forms::ImageLayout::None;
			this->button1->Cursor = System::Windows::Forms::Cursors::Default;
			this->button1->FlatAppearance->BorderSize = 0;
			this->button1->FlatAppearance->MouseDownBackColor = System::Drawing::Color::Transparent;
			this->button1->FlatAppearance->MouseOverBackColor = System::Drawing::Color::Transparent;
			this->button1->FlatStyle = System::Windows::Forms::FlatStyle::Flat;
			this->button1->Font = (gcnew System::Drawing::Font(L"Century Gothic", 18, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->button1->ForeColor = System::Drawing::Color::Black;
			this->button1->Image = (cli::safe_cast<System::Drawing::Image^>(resources->GetObject(L"button1.Image")));
			this->button1->Location = System::Drawing::Point(118, 568);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(326, 320);
			this->button1->TabIndex = 1;
			this->button1->TabStop = false;
			this->button1->Text = L"New Volunteer ";
			this->button1->UseVisualStyleBackColor = false;
			this->button1->Click += gcnew System::EventHandler(this, &MyForm::button1_Click);
			// 
			// button2
			// 
			this->button2->AutoSize = true;
			this->button2->BackColor = System::Drawing::Color::Transparent;
			this->button2->FlatAppearance->BorderSize = 0;
			this->button2->FlatAppearance->MouseDownBackColor = System::Drawing::Color::Transparent;
			this->button2->FlatAppearance->MouseOverBackColor = System::Drawing::Color::Transparent;
			this->button2->FlatStyle = System::Windows::Forms::FlatStyle::Flat;
			this->button2->Font = (gcnew System::Drawing::Font(L"Century Gothic", 18, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->button2->ForeColor = System::Drawing::Color::Black;
			this->button2->Image = (cli::safe_cast<System::Drawing::Image^>(resources->GetObject(L"button2.Image")));
			this->button2->Location = System::Drawing::Point(793, 568);
			this->button2->Name = L"button2";
			this->button2->Size = System::Drawing::Size(370, 320);
			this->button2->TabIndex = 2;
			this->button2->TabStop = false;
			this->button2->Text = L"Registered Volunteer";
			this->button2->UseVisualStyleBackColor = false;
			this->button2->Click += gcnew System::EventHandler(this, &MyForm::button2_Click);
			// 
			// fileToolStripMenuItem
			// 
			this->fileToolStripMenuItem->BackColor = System::Drawing::Color::Transparent;
			this->fileToolStripMenuItem->DropDownItems->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(2) {
				this->newLogToolStripMenuItem,
					this->exitToolStripMenuItem
			});
			this->fileToolStripMenuItem->Font = (gcnew System::Drawing::Font(L"Segoe UI", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->fileToolStripMenuItem->ForeColor = System::Drawing::Color::White;
			this->fileToolStripMenuItem->Name = L"fileToolStripMenuItem";
			this->fileToolStripMenuItem->Size = System::Drawing::Size(46, 25);
			this->fileToolStripMenuItem->Text = L"File";
			// 
			// newLogToolStripMenuItem
			// 
			this->newLogToolStripMenuItem->Name = L"newLogToolStripMenuItem";
			this->newLogToolStripMenuItem->Size = System::Drawing::Size(142, 26);
			this->newLogToolStripMenuItem->Text = L"New Log";
			this->newLogToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::newLogToolStripMenuItem_Click);
			// 
			// exitToolStripMenuItem
			// 
			this->exitToolStripMenuItem->Name = L"exitToolStripMenuItem";
			this->exitToolStripMenuItem->Size = System::Drawing::Size(142, 26);
			this->exitToolStripMenuItem->Text = L"Exit";
			this->exitToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::exitToolStripMenuItem_Click);
			// 
			// menuStrip1
			// 
			this->menuStrip1->BackColor = System::Drawing::Color::Transparent;
			this->menuStrip1->Items->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(2) {
				this->fileToolStripMenuItem,
					this->editToolStripMenuItem
			});
			this->menuStrip1->Location = System::Drawing::Point(0, 0);
			this->menuStrip1->Name = L"menuStrip1";
			this->menuStrip1->Size = System::Drawing::Size(1331, 29);
			this->menuStrip1->TabIndex = 0;
			this->menuStrip1->Text = L"menuStrip1";
			// 
			// editToolStripMenuItem
			// 
			this->editToolStripMenuItem->BackColor = System::Drawing::Color::Transparent;
			this->editToolStripMenuItem->DropDownItems->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(1) { this->memberInfoToolStripMenuItem });
			this->editToolStripMenuItem->Font = (gcnew System::Drawing::Font(L"Segoe UI", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->editToolStripMenuItem->ForeColor = System::Drawing::Color::White;
			this->editToolStripMenuItem->Name = L"editToolStripMenuItem";
			this->editToolStripMenuItem->Size = System::Drawing::Size(48, 25);
			this->editToolStripMenuItem->Text = L"Edit";
			// 
			// memberInfoToolStripMenuItem
			// 
			this->memberInfoToolStripMenuItem->Name = L"memberInfoToolStripMenuItem";
			this->memberInfoToolStripMenuItem->Size = System::Drawing::Size(170, 26);
			this->memberInfoToolStripMenuItem->Text = L"Member Info";
			this->memberInfoToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::memberInfoToolStripMenuItem_Click);
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->BackColor = System::Drawing::Color::Transparent;
			this->label1->Font = (gcnew System::Drawing::Font(L"Century Gothic", 36, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->label1->ForeColor = System::Drawing::Color::Plum;
			this->label1->Location = System::Drawing::Point(392, 9);
			this->label1->Name = L"label1";
			this->label1->Padding = System::Windows::Forms::Padding(20);
			this->label1->Size = System::Drawing::Size(454, 98);
			this->label1->TabIndex = 7;
			this->label1->Text = L"Volunteer Sign-In";
			// 
			// pictureBox2
			// 
			this->pictureBox2->BackColor = System::Drawing::Color::Transparent;
			this->pictureBox2->Image = (cli::safe_cast<System::Drawing::Image^>(resources->GetObject(L"pictureBox2.Image")));
			this->pictureBox2->Location = System::Drawing::Point(108, 380);
			this->pictureBox2->Name = L"pictureBox2";
			this->pictureBox2->Size = System::Drawing::Size(1055, 218);
			this->pictureBox2->SizeMode = System::Windows::Forms::PictureBoxSizeMode::Zoom;
			this->pictureBox2->TabIndex = 9;
			this->pictureBox2->TabStop = false;
			// 
			// pictureBox1
			// 
			this->pictureBox1->BackColor = System::Drawing::Color::Transparent;
			this->pictureBox1->Image = (cli::safe_cast<System::Drawing::Image^>(resources->GetObject(L"pictureBox1.Image")));
			this->pictureBox1->Location = System::Drawing::Point(414, 136);
			this->pictureBox1->Name = L"pictureBox1";
			this->pictureBox1->Size = System::Drawing::Size(432, 238);
			this->pictureBox1->SizeMode = System::Windows::Forms::PictureBoxSizeMode::StretchImage;
			this->pictureBox1->TabIndex = 10;
			this->pictureBox1->TabStop = false;
			// 
			// panel1
			// 
			this->panel1->BackColor = System::Drawing::Color::Transparent;
			this->panel1->Controls->Add(this->panel5);
			this->panel1->Location = System::Drawing::Point(87, 571);
			this->panel1->Name = L"panel1";
			this->panel1->Size = System::Drawing::Size(61, 320);
			this->panel1->TabIndex = 11;
			// 
			// panel5
			// 
			this->panel5->BackColor = System::Drawing::Color::Transparent;
			this->panel5->Location = System::Drawing::Point(11, 288);
			this->panel5->Name = L"panel5";
			this->panel5->Size = System::Drawing::Size(1065, 32);
			this->panel5->TabIndex = 15;
			// 
			// panel2
			// 
			this->panel2->BackColor = System::Drawing::Color::Transparent;
			this->panel2->Location = System::Drawing::Point(414, 597);
			this->panel2->Name = L"panel2";
			this->panel2->Size = System::Drawing::Size(30, 294);
			this->panel2->TabIndex = 12;
			// 
			// panel3
			// 
			this->panel3->BackColor = System::Drawing::Color::Transparent;
			this->panel3->Location = System::Drawing::Point(782, 597);
			this->panel3->Name = L"panel3";
			this->panel3->Size = System::Drawing::Size(30, 294);
			this->panel3->TabIndex = 13;
			// 
			// panel4
			// 
			this->panel4->BackColor = System::Drawing::Color::Transparent;
			this->panel4->Location = System::Drawing::Point(1144, 597);
			this->panel4->Name = L"panel4";
			this->panel4->Size = System::Drawing::Size(28, 294);
			this->panel4->TabIndex = 14;
			// 
			// panel6
			// 
			this->panel6->BackColor = System::Drawing::Color::Transparent;
			this->panel6->Controls->Add(this->panel7);
			this->panel6->Location = System::Drawing::Point(46, 859);
			this->panel6->Name = L"panel6";
			this->panel6->Size = System::Drawing::Size(1233, 42);
			this->panel6->TabIndex = 16;
			// 
			// panel7
			// 
			this->panel7->BackColor = System::Drawing::Color::Transparent;
			this->panel7->Location = System::Drawing::Point(11, 288);
			this->panel7->Name = L"panel7";
			this->panel7->Size = System::Drawing::Size(1065, 32);
			this->panel7->TabIndex = 15;
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->AutoScroll = true;
			this->BackColor = System::Drawing::Color::White;
			this->BackgroundImage = (cli::safe_cast<System::Drawing::Image^>(resources->GetObject(L"$this.BackgroundImage")));
			this->BackgroundImageLayout = System::Windows::Forms::ImageLayout::Stretch;
			this->ClientSize = System::Drawing::Size(1331, 909);
			this->Controls->Add(this->panel6);
			this->Controls->Add(this->panel4);
			this->Controls->Add(this->panel3);
			this->Controls->Add(this->panel1);
			this->Controls->Add(this->pictureBox2);
			this->Controls->Add(this->menuStrip1);
			this->Controls->Add(this->label1);
			this->Controls->Add(this->button2);
			this->Controls->Add(this->pictureBox1);
			this->Controls->Add(this->panel2);
			this->Controls->Add(this->button1);
			this->ForeColor = System::Drawing::Color::CornflowerBlue;
			this->Icon = (cli::safe_cast<System::Drawing::Icon^>(resources->GetObject(L"$this.Icon")));
			this->MainMenuStrip = this->menuStrip1;
			this->Name = L"MyForm";
			this->Text = L"MyForm";
			this->menuStrip1->ResumeLayout(false);
			this->menuStrip1->PerformLayout();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox2))->EndInit();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox1))->EndInit();
			this->panel1->ResumeLayout(false);
			this->panel6->ResumeLayout(false);
			this->ResumeLayout(false);
			this->PerformLayout();

		}
	


#pragma endregion

	private: System::Void newLogToolStripMenuItem_Click(System::Object^  sender, System::EventArgs^  e) {



		////NetOffice::ExcelApi::Range^ cell = oSheet->Range(oSheet->Cells[1, 1], oSheet->Cells[1000, 1000]);    ******* Compatibility Error in Excel 2000 (can't create range object)
		////cell->ColumnWidth = 10;																				 ******* Compatibility Error in Excel 2000


		NetOffice::ExcelApi::Application^ oApp;
		NetOffice::ExcelApi::Worksheet^ oSheet;
		NetOffice::ExcelApi::Workbook^ oBook;

		oApp = gcnew NetOffice::ExcelApi::Application();
		oBook = oApp->Workbooks->Add();
		oSheet = (NetOffice::ExcelApi::Worksheet^)oBook->Worksheets[1];

		oApp->DisplayAlerts = false;

		for (int k = 1; k < 20; k++) {
			oSheet->Columns[k]->ColumnWidth = 10;
		}


		oSheet->Range(oSheet->Cells[1, 1], oSheet->Cells[1, 3])->Merge(Type::Missing);	//Merges two cells
		oSheet->Cells[1, 1]->Value = "Volunteer Names";
		oSheet->Cells[1, 7]->Value = "Count:";
		oSheet->Cells[1, 8]->Value = "0";



		oBook->SaveAs("C:\\Volunteer Log");
		oApp->Quit();
		oApp->DisposeChildInstances();

		MessageBox::Show(Owner, "New Log Created In C Drive", "Volunteer Log", MessageBoxButtons::OK, MessageBoxIcon::Information);
		delete oApp;
	}


	private: System::Void button1_Click(System::Object^  sender, System::EventArgs^  e) {

		button1->FlatAppearance->BorderSize = 0;


		char str[] = "C:\\Volunteer Log.xlsx";
		char str2[] = "C:\\Volunteer Log.xls";
		char * str1 = str;
		struct stat buf;
		struct stat buf2;
		int x = 2;
		x = stat(str, &buf);
		if (stat(str, &buf) == 0 || stat(str2, &buf2) == 0) {

			DataBox^ b = gcnew DataBox;
			b->Show();
		}
		else {

			MessageBox::Show(Owner, "No log found. Create a new log. (File -> New Log)", "Volunteer Log", MessageBoxButtons::OK, MessageBoxIcon::Information);

		}
	}







	private: System::Void exitToolStripMenuItem_Click(System::Object^  sender, System::EventArgs^  e) {
		this->Close();
	}


	private: System::Void button2_Click(System::Object^  sender, System::EventArgs^  e) {




		_Application^ oApp;
		_Worksheet^ oSheet;
		_Workbook^ oBook;


		oApp = gcnew NetOffice::ExcelApi::Application();
		oBook = oApp->Workbooks->Open("C:\\Volunteer Log");
		oSheet = (NetOffice::ExcelApi::Worksheet^) oBook->ActiveSheet;



		System::String^ str = oSheet->Cells[1, 8]->Value->ToString();
		int i = int::Parse(str);

		int max = i * 6;


		List^ list = gcnew List;
		for (int i = 0; i < max; i = i + 6) {

			list->listBox1->Items->Add((System::Object^) oSheet->Cells[i + 2, 1]->Value);


		}
		list->Show();


		oBook->Save();
		oApp->Quit();
		oApp->DisposeChildInstances();
		delete oApp;


	}



private: System::Void memberInfoToolStripMenuItem_Click(System::Object^  sender, System::EventArgs^  e) {

	_Application^ oApp;
	_Worksheet^ oSheet;
	_Workbook^ oBook;


	oApp = gcnew NetOffice::ExcelApi::Application();
	oBook = oApp->Workbooks->Open("C:\\Volunteer Log");
	oSheet = (NetOffice::ExcelApi::Worksheet^) oBook->ActiveSheet;



	System::String^ str = oSheet->Cells[1, 8]->Value->ToString();
	int i = int::Parse(str);

	int max = i * 6;

	listEdit^ b = gcnew listEdit;
	for (int i = 0; i < max; i = i + 6) {

		b->listBox1->Items->Add((System::Object^) oSheet->Cells[i + 2, 1]->Value);


	}
	b->Show();


	oBook->Save();
	oApp->Quit();
	oApp->DisposeChildInstances();
	delete oApp;
}


};

}
