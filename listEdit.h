/****************************************************************************************
Populates the List of Registered Volunteered that can be Edited
*****************************************************************************************/
#pragma once
#include "editInfo.h"

namespace Data_Base_2 {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// Summary for listEdit
	/// </summary>
	public ref class listEdit : public System::Windows::Forms::Form
	{
	public:
		listEdit(void)
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
		~listEdit()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Label^  label1;
	public: System::Windows::Forms::ListBox^  listBox1;
	private: System::Windows::Forms::Button^  button1;
	private: System::Windows::Forms::PictureBox^  pictureBox1;
	public:
	private:
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
			System::ComponentModel::ComponentResourceManager^  resources = (gcnew System::ComponentModel::ComponentResourceManager(listEdit::typeid));
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->listBox1 = (gcnew System::Windows::Forms::ListBox());
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->pictureBox1 = (gcnew System::Windows::Forms::PictureBox());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox1))->BeginInit();
			this->SuspendLayout();
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->BackColor = System::Drawing::Color::White;
			this->label1->Font = (gcnew System::Drawing::Font(L"Century Gothic", 20.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->label1->Location = System::Drawing::Point(175, 44);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(344, 33);
			this->label1->TabIndex = 2;
			this->label1->Text = L"Please Select Your Name";
			// 
			// listBox1
			// 
			this->listBox1->Font = (gcnew System::Drawing::Font(L"Calibri", 18, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->listBox1->FormattingEnabled = true;
			this->listBox1->ItemHeight = 29;
			this->listBox1->Location = System::Drawing::Point(88, 106);
			this->listBox1->Name = L"listBox1";
			this->listBox1->Size = System::Drawing::Size(522, 178);
			this->listBox1->TabIndex = 3;
			// 
			// button1
			// 
			this->button1->Font = (gcnew System::Drawing::Font(L"Century Gothic", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->button1->Location = System::Drawing::Point(251, 313);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(182, 32);
			this->button1->TabIndex = 4;
			this->button1->Text = L"Edit Member";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Click += gcnew System::EventHandler(this, &listEdit::button1_Click);
			// 
			// pictureBox1
			// 
			this->pictureBox1->BackColor = System::Drawing::Color::Transparent;
			this->pictureBox1->Image = (cli::safe_cast<System::Drawing::Image^>(resources->GetObject(L"pictureBox1.Image")));
			this->pictureBox1->Location = System::Drawing::Point(583, 290);
			this->pictureBox1->Name = L"pictureBox1";
			this->pictureBox1->Size = System::Drawing::Size(113, 88);
			this->pictureBox1->SizeMode = System::Windows::Forms::PictureBoxSizeMode::StretchImage;
			this->pictureBox1->TabIndex = 12;
			this->pictureBox1->TabStop = false;
			// 
			// listEdit
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->BackColor = System::Drawing::Color::White;
			this->ClientSize = System::Drawing::Size(695, 382);
			this->Controls->Add(this->pictureBox1);
			this->Controls->Add(this->button1);
			this->Controls->Add(this->listBox1);
			this->Controls->Add(this->label1);
			this->Name = L"listEdit";
			this->Text = L"listEdit";
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox1))->EndInit();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion

	private: System::Void button1_Click(System::Object^  sender, System::EventArgs^  e) {

		if (listBox1->SelectedItem == nullptr) {
			MessageBox::Show(Owner, "Please select a name.", "Volunteer Log", MessageBoxButtons::OK, MessageBoxIcon::Information);
		
		}

		else {
			int i = listBox1->SelectedIndex;
			i = i * 6;

			NetOffice::ExcelApi::Application^ oApp;
			NetOffice::ExcelApi::Worksheet^ oSheet;
			NetOffice::ExcelApi::Workbook^ oBook;

			oApp = gcnew NetOffice::ExcelApi::Application();
			oBook = oApp->Workbooks->Open("C:\\Volunteer Log");
			oSheet = (NetOffice::ExcelApi::Worksheet^) oBook->ActiveSheet;

			editInfo^ b = gcnew editInfo;

			b->name->Text = oSheet->Cells[i + 2, 1]->Value->ToString();



			if (oSheet->Cells[i + 4, 1]->Value != nullptr) {
				b->email->Text = oSheet->Cells[i + 4, 1]->Value->ToString();
			}

			if (oSheet->Cells[i + 3, 1]->Value != nullptr) {

				b->phone->Text = oSheet->Cells[i + 3, 1]->Value->ToString();
			}




			editInfo::nameIndex = i;

			oBook->Save();
			oApp->Quit();
			oApp->DisposeChildInstances();
			delete oApp;







			b->Show();
			this->Close();
		}
	}
	};
}
