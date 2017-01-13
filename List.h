/****************************************************************************************
Registered Volunteer Selection Form
*****************************************************************************************/
#pragma once
#include "Date.h"
#include "Information.h"
#include <string>

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
	/// Summary for List
	/// </summary>
	public ref class List : public System::Windows::Forms::Form
	{
	public:
		List(void)
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
		~List()
		{
			if (components)
			{
				delete components;
			}
		}
	public: System::Windows::Forms::ListBox^  listBox1;
	private: System::Windows::Forms::Label^  label1;
	public:
	private: System::Windows::Forms::Button^  button1;
	private: System::Windows::Forms::Button^  button2;
	private: System::Windows::Forms::PictureBox^  pictureBox1;
	protected:

	public:
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
			System::ComponentModel::ComponentResourceManager^  resources = (gcnew System::ComponentModel::ComponentResourceManager(List::typeid));
			this->listBox1 = (gcnew System::Windows::Forms::ListBox());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->button2 = (gcnew System::Windows::Forms::Button());
			this->pictureBox1 = (gcnew System::Windows::Forms::PictureBox());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox1))->BeginInit();
			this->SuspendLayout();
			// 
			// listBox1
			// 
			this->listBox1->Font = (gcnew System::Drawing::Font(L"Calibri", 18, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->listBox1->FormattingEnabled = true;
			this->listBox1->ItemHeight = 29;
			this->listBox1->Location = System::Drawing::Point(73, 84);
			this->listBox1->Name = L"listBox1";
			this->listBox1->Size = System::Drawing::Size(522, 178);
			this->listBox1->TabIndex = 0;
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->BackColor = System::Drawing::Color::White;
			this->label1->Font = (gcnew System::Drawing::Font(L"Century Gothic", 20.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->label1->Location = System::Drawing::Point(172, 21);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(344, 33);
			this->label1->TabIndex = 1;
			this->label1->Text = L"Please Select Your Name";
			// 
			// button1
			// 
			this->button1->Font = (gcnew System::Drawing::Font(L"Century Gothic", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->button1->Location = System::Drawing::Point(112, 293);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(108, 30);
			this->button1->TabIndex = 2;
			this->button1->Text = L"Select";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Click += gcnew System::EventHandler(this, &List::button1_Click);
			// 
			// button2
			// 
			this->button2->Font = (gcnew System::Drawing::Font(L"Century Gothic", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->button2->Location = System::Drawing::Point(347, 293);
			this->button2->Name = L"button2";
			this->button2->Size = System::Drawing::Size(197, 30);
			this->button2->TabIndex = 3;
			this->button2->Text = L"View Volunteer History";
			this->button2->UseVisualStyleBackColor = true;
			this->button2->Click += gcnew System::EventHandler(this, &List::button2_Click);
			// 
			// pictureBox1
			// 
			this->pictureBox1->BackColor = System::Drawing::Color::Transparent;
			this->pictureBox1->Image = (cli::safe_cast<System::Drawing::Image^>(resources->GetObject(L"pictureBox1.Image")));
			this->pictureBox1->Location = System::Drawing::Point(564, 278);
			this->pictureBox1->Name = L"pictureBox1";
			this->pictureBox1->Size = System::Drawing::Size(113, 88);
			this->pictureBox1->SizeMode = System::Windows::Forms::PictureBoxSizeMode::StretchImage;
			this->pictureBox1->TabIndex = 12;
			this->pictureBox1->TabStop = false;
			// 
			// List
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->BackColor = System::Drawing::Color::White;
			this->ClientSize = System::Drawing::Size(677, 367);
			this->Controls->Add(this->pictureBox1);
			this->Controls->Add(this->button2);
			this->Controls->Add(this->button1);
			this->Controls->Add(this->label1);
			this->Controls->Add(this->listBox1);
			this->Name = L"List";
			this->Text = L"List";
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox1))->EndInit();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
	private: System::Void button1_Click(System::Object^  sender, System::EventArgs^  e) {
		if (listBox1->SelectedItem != nullptr) {
			int i = listBox1->SelectedIndex;
			i = i * 6;

			Date^ b = gcnew Date;
			b->Show();
			b->namePosition = i;

			this->Close();
		}
		else {
			MessageBox::Show(Owner, "Please select your name.", "Volunteer Log", MessageBoxButtons::OK, MessageBoxIcon::Information);

		}
	}
	private: System::Void button2_Click(System::Object^  sender, System::EventArgs^  e) {

		if (listBox1->SelectedItem != nullptr) {

			int i = listBox1->SelectedIndex;
			i = i * 6;

			_Application^ oApp;
			_Worksheet^ oSheet;
			_Workbook^ oBook;

			oApp = gcnew NetOffice::ExcelApi::Application();
			oBook = oApp->Workbooks->Open("C:\\Volunteer Log");
			oSheet = (NetOffice::ExcelApi::Worksheet^) oBook->ActiveSheet;


			Information^ b = gcnew Information;
			b->label9->Text = oSheet->Cells[i + 2, 1]->Value->ToString();												//NAME

			System::String^ strx = oSheet->Cells[i + 2, 5]->Value->ToString();
			strx = strx->Remove(strx->Length - 12, 12);
			b->label6->Text = strx;																						//First Date

			strx = oSheet->Cells[i + 5, 1]->Value->ToString();
			strx = strx->Substring(6, strx->Length - 6);

			b->label7->Text = strx;																						//Total hours

			System::String^ str = oSheet->Cells[i + 5, 1]->Value->ToString();
			str = str->Remove(0, 7);
			i = int::Parse(str);

			if (i >= 2500) {
				b->label8->Text = "4 (QUARKS)";
			}
			else if (i >= 1000) {
				b->label8->Text = "4 (Fermium)";
			}
			else if (i >= 700) {
				b->label8->Text = "3 (Curium)";
			}
			else if (i >= 400) {
				b->label8->Text = "3 (Gold)";
			}
			else if (i >= 250) {
				b->label8->Text = "3 (Platinum)";
			}
			else if (i >= 100) {
				b->label8->Text = "3 (Promethium)";
			}
			else if (i >= 60) {
				b->label8->Text = "2 (Xenon)";
			}
			else if (i >= 40) {
				b->label8->Text = "2 (Krypton)";
			}
			else if (i >= 20) {
				b->label8->Text = "2 (Argon)";
			}
			else if (i >= 10) {
				b->label8->Text = "2 (Neon)";
			}
			else {
				b->label8->Text = "1 (Hydrogen)";
			}

			this->Close();


			b->Show();






			oBook->Save();
			oApp->Quit();
			oApp->DisposeChildInstances();
			delete oApp;

		}

		else {

			MessageBox::Show(Owner, "Please select your name.", "Volunteer Log", MessageBoxButtons::OK, MessageBoxIcon::Information);
		}
	}


};
}
