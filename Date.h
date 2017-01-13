#pragma once
#include <msclr\marshal_cppstd.h>
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
	/// Summary for Date
	/// </summary>
	public ref class Date : public System::Windows::Forms::Form
	{
	public:
		Date(void)
		{
			InitializeComponent();
			//
			//TODO: Add the constructor code here
			//
		}
	private: System::Windows::Forms::PictureBox^  pictureBox1;
	public:

		static int namePosition;

	protected:
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~Date()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Label^  label1;
	protected:
	private: System::Windows::Forms::Label^  label2;
	private: System::Windows::Forms::DateTimePicker^  dateTimePicker1;
	private: System::Windows::Forms::Label^  label3;





	private: System::Windows::Forms::Label^  label4;


	private: System::Windows::Forms::Button^  button1;
	private: System::Windows::Forms::CheckedListBox^  checkedListBox1;
	private: System::Windows::Forms::CheckedListBox^  checkedListBox2;


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
			System::ComponentModel::ComponentResourceManager^  resources = (gcnew System::ComponentModel::ComponentResourceManager(Date::typeid));
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->label2 = (gcnew System::Windows::Forms::Label());
			this->dateTimePicker1 = (gcnew System::Windows::Forms::DateTimePicker());
			this->label3 = (gcnew System::Windows::Forms::Label());
			this->label4 = (gcnew System::Windows::Forms::Label());
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->checkedListBox1 = (gcnew System::Windows::Forms::CheckedListBox());
			this->checkedListBox2 = (gcnew System::Windows::Forms::CheckedListBox());
			this->pictureBox1 = (gcnew System::Windows::Forms::PictureBox());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox1))->BeginInit();
			this->SuspendLayout();
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->Font = (gcnew System::Drawing::Font(L"Calibri", 14.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->label1->Location = System::Drawing::Point(16, 100);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(108, 23);
			this->label1->TabIndex = 0;
			this->label1->Text = L"Sign In Time:";
			// 
			// label2
			// 
			this->label2->AutoSize = true;
			this->label2->Font = (gcnew System::Drawing::Font(L"Calibri", 14.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->label2->Location = System::Drawing::Point(351, 100);
			this->label2->Name = L"label2";
			this->label2->Size = System::Drawing::Size(121, 23);
			this->label2->TabIndex = 1;
			this->label2->Text = L"Sign Out Time:";
			// 
			// dateTimePicker1
			// 
			this->dateTimePicker1->Font = (gcnew System::Drawing::Font(L"Arial Rounded MT Bold", 14.25F, System::Drawing::FontStyle::Regular,
				System::Drawing::GraphicsUnit::Point, static_cast<System::Byte>(0)));
			this->dateTimePicker1->Location = System::Drawing::Point(130, 26);
			this->dateTimePicker1->Name = L"dateTimePicker1";
			this->dateTimePicker1->Size = System::Drawing::Size(342, 29);
			this->dateTimePicker1->TabIndex = 2;
			// 
			// label3
			// 
			this->label3->AutoSize = true;
			this->label3->Font = (gcnew System::Drawing::Font(L"Calibri", 14.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->label3->Location = System::Drawing::Point(72, 26);
			this->label3->Name = L"label3";
			this->label3->Size = System::Drawing::Size(45, 23);
			this->label3->TabIndex = 3;
			this->label3->Text = L"Day:";
			// 
			// label4
			// 
			this->label4->AutoSize = true;
			this->label4->Location = System::Drawing::Point(234, 182);
			this->label4->Name = L"label4";
			this->label4->Size = System::Drawing::Size(10, 13);
			this->label4->TabIndex = 9;
			this->label4->Text = L":";
			// 
			// button1
			// 
			this->button1->BackColor = System::Drawing::Color::White;
			this->button1->FlatAppearance->BorderColor = System::Drawing::Color::Black;
			this->button1->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->button1->Location = System::Drawing::Point(273, 281);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(112, 36);
			this->button1->TabIndex = 12;
			this->button1->Text = L"Done";
			this->button1->UseVisualStyleBackColor = false;
			this->button1->Click += gcnew System::EventHandler(this, &Date::button1_Click);
			this->button1->MouseEnter += gcnew System::EventHandler(this, &Date::button1_MouseEnter);
			this->button1->MouseLeave += gcnew System::EventHandler(this, &Date::button1_MouseLeave);
			// 
			// checkedListBox1
			// 
			this->checkedListBox1->BackColor = System::Drawing::Color::Aqua;
			this->checkedListBox1->CheckOnClick = true;
			this->checkedListBox1->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 11.25F, System::Drawing::FontStyle::Regular,
				System::Drawing::GraphicsUnit::Point, static_cast<System::Byte>(0)));
			this->checkedListBox1->FormattingEnabled = true;
			this->checkedListBox1->Items->AddRange(gcnew cli::array< System::Object^  >(20) {
				L"5:00 AM", L"6:00 AM", L"7:00 AM", L"8:00 AM",
					L"9:00 AM", L"10:00 AM", L"11:00 AM", L"12:00 PM", L"1:00 PM", L"2:00 PM", L"3:00 PM", L"4:00 PM", L"5:00 PM", L"6:00 PM", L"7:00 PM",
					L"8:00 PM", L"9:00 PM", L"10:00 PM", L"11:00 PM", L"12:00 AM"
			});
			this->checkedListBox1->Location = System::Drawing::Point(130, 100);
			this->checkedListBox1->Name = L"checkedListBox1";
			this->checkedListBox1->Size = System::Drawing::Size(107, 137);
			this->checkedListBox1->TabIndex = 13;
			this->checkedListBox1->ItemCheck += gcnew System::Windows::Forms::ItemCheckEventHandler(this, &Date::checkedListBox1_ItemCheck);
			// 
			// checkedListBox2
			// 
			this->checkedListBox2->BackColor = System::Drawing::Color::Aqua;
			this->checkedListBox2->CheckOnClick = true;
			this->checkedListBox2->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 11.25F, System::Drawing::FontStyle::Regular,
				System::Drawing::GraphicsUnit::Point, static_cast<System::Byte>(0)));
			this->checkedListBox2->FormattingEnabled = true;
			this->checkedListBox2->Items->AddRange(gcnew cli::array< System::Object^  >(20) {
				L"5:00 AM", L"6:00 AM", L"7:00 AM", L"8:00 AM",
					L"9:00 AM", L"10:00 AM", L"11:00 AM", L"12:00 PM", L"1:00 PM", L"2:00 PM", L"3:00 PM", L"4:00 PM", L"5:00 PM", L"6:00 PM", L"7:00 PM",
					L"8:00 PM", L"9:00 PM", L"10:00 PM", L"11:00 PM", L"12:00 AM"
			});
			this->checkedListBox2->Location = System::Drawing::Point(474, 100);
			this->checkedListBox2->Name = L"checkedListBox2";
			this->checkedListBox2->Size = System::Drawing::Size(116, 137);
			this->checkedListBox2->TabIndex = 14;
			this->checkedListBox2->ItemCheck += gcnew System::Windows::Forms::ItemCheckEventHandler(this, &Date::checkedListBox2_ItemCheck);
			// 
			// pictureBox1
			// 
			this->pictureBox1->BackColor = System::Drawing::Color::Transparent;
			this->pictureBox1->Image = (cli::safe_cast<System::Drawing::Image^>(resources->GetObject(L"pictureBox1.Image")));
			this->pictureBox1->Location = System::Drawing::Point(537, 281);
			this->pictureBox1->Name = L"pictureBox1";
			this->pictureBox1->Size = System::Drawing::Size(113, 88);
			this->pictureBox1->SizeMode = System::Windows::Forms::PictureBoxSizeMode::StretchImage;
			this->pictureBox1->TabIndex = 15;
			this->pictureBox1->TabStop = false;
			// 
			// Date
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->BackColor = System::Drawing::Color::Orange;
			this->ClientSize = System::Drawing::Size(649, 370);
			this->Controls->Add(this->pictureBox1);
			this->Controls->Add(this->checkedListBox2);
			this->Controls->Add(this->checkedListBox1);
			this->Controls->Add(this->button1);
			this->Controls->Add(this->label4);
			this->Controls->Add(this->label3);
			this->Controls->Add(this->dateTimePicker1);
			this->Controls->Add(this->label2);
			this->Controls->Add(this->label1);
			this->Name = L"Date";
			this->Text = L"Date";
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox1))->EndInit();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion

private: System::Void listBox1_SelectedIndexChanged(System::Object^  sender, System::EventArgs^  e) {
}

private: System::Void button1_Click(System::Object^  sender, System::EventArgs^  e) {

	System::String^ strx=nullptr;

	if (checkedListBox1->SelectedItem == nullptr) {
		MessageBox::Show(Owner, "Please select a check-in time.", "Volunteer Log", MessageBoxButtons::OK, MessageBoxIcon::Information);
	}
	else if (checkedListBox2->SelectedItem == nullptr) {
		MessageBox::Show(Owner, "Please select a check-out time.", "Volunteer Log", MessageBoxButtons::OK, MessageBoxIcon::Information);

	}

	else {
		int i = namePosition;
		int j = 5;

		_Application^ oApp;
		_Worksheet^ oSheet;
		_Workbook^ oBook;

		oApp = gcnew NetOffice::ExcelApi::Application();
		oBook = oApp->Workbooks->Open("C:\\Volunteer Log");
		oSheet = (NetOffice::ExcelApi::Worksheet^) oBook->ActiveSheet;


		System::String^ month = dateTimePicker1->Value.Date.Date.Month.ToString();
		System::String^ day = dateTimePicker1->Value.Date.Date.Day.ToString();
		System::String^ year = dateTimePicker1->Value.Date.Date.Year.ToString();

		System::String^ date = month + "/" + day + "/" + year;

		if (oSheet->Cells[i + 2, j]->Value != nullptr) {
			strx = oSheet->Cells[i + 2, j]->Value->ToString();
			strx = strx->Remove(strx->Length - 12, 12);
		}


		while (oSheet->Cells[i + 2, j]->Value != nullptr && strx != date) {													
			j++;

			if (oSheet->Cells[i + 2, j]->Value != nullptr) {
				strx = oSheet->Cells[i + 2, j]->Value->ToString();
				strx = strx->Remove(strx->Length - 12, 12);
			}	
			 
		}																						
		oSheet->Cells[i + 2, j]->Value = date;																										//Date



		oSheet->Cells[i + 3, j]->Value = checkedListBox1->SelectedItem->ToString();																	//Time IN
		std::string timeIn = msclr::interop::marshal_as <std::string>(checkedListBox1->SelectedItem->ToString());







		oSheet->Cells[i + 4, j]->Value = checkedListBox2->SelectedItem->ToString();																	//Time OUT
		std::string timeOut = msclr::interop::marshal_as <std::string>(checkedListBox2->SelectedItem->ToString());




		if (timeIn[1] == '0' || timeIn[1] == '1' || timeIn[1] == '2') {
			timeIn = timeIn.substr(0, 2);
		}
		else {
			timeIn = timeIn.substr(0, 1);
		}
		int time1 = std::stoi(timeIn);

		if (timeOut[1] == '0' || timeOut[1] == '1' || timeOut[1] == '2') {
			timeOut = timeOut.substr(0, 2);
		}
		else {
			timeOut = timeOut.substr(0, 1);
		}
		int time2 = std::stoi(timeOut);




		if (time2 - time1 < 0) {
			oSheet->Cells[i + 5, j]->Value = (12 - time1) + time2;
		}
		else if (time2 - time1 == 0) {
			oSheet->Cells[i + 5, j]->Value = 12;
		}
		else {
			oSheet->Cells[i + 5, j]->Value = time2 - time1;
		}

		j = 5;

		int count = 0;
		while (oSheet->Cells[i + 5, j]->Value != nullptr) {																								//Hours
			std::string total = msclr::interop::marshal_as<std::string>(oSheet->Cells[i + 5, j]->Value->ToString());
			count = count + stoi(total);
			j++;
		}

		oSheet->Cells[i + 5, 1]->Value = "Total: " + count;




		oBook->Save();
		oApp->Quit();
		this->Close();
		oApp->DisposeChildInstances();
		delete oApp;

		MessageBox::Show(Owner, "Information Saved!", "Volunteer Log", MessageBoxButtons::OK, MessageBoxIcon::Information);


	}

}

private: System::Void checkedListBox1_ItemCheck(System::Object^  sender, System::Windows::Forms::ItemCheckEventArgs^  e) {
	for (int ix = 0; ix < checkedListBox1->Items->Count; ++ix)
		if (ix != e->Index) checkedListBox1->SetItemChecked(ix, false);
}

private: System::Void checkedListBox2_ItemCheck(System::Object^  sender, System::Windows::Forms::ItemCheckEventArgs^  e) {
	for (int ix = 0; ix < checkedListBox2->Items->Count; ++ix)
		if (ix != e->Index) checkedListBox2->SetItemChecked(ix, false);
}







private: System::Void button1_MouseEnter(System::Object^  sender, System::EventArgs^  e) {
	button1->BackColor = System::Drawing::Color::LightGreen;
}
private: System::Void button1_MouseLeave(System::Object^  sender, System::EventArgs^  e) {
	button1->BackColor = System::Drawing::Color::White;
}

};
}


