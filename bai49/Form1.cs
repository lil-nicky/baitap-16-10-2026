
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ClinicAppointmentManager
{
    public class Appointment
    {
        public string AppointmentId { get; set; }
        public string PatientName { get; set; }
        public string Phone { get; set; }
        public DateTime AppointmentDate { get; set; }
        public string Doctor { get; set; }
        public string Reason { get; set; }
        public string Status { get; set; }
    }

    public partial class Form1 : Form
    {
        private readonly List<Appointment> appointments =
            new List<Appointment>();

        private readonly BindingList<Appointment> displayedAppointments =
            new BindingList<Appointment>();

        private readonly BindingSource bindingSource =
            new BindingSource();

        private readonly ErrorProvider errorProvider =
            new ErrorProvider();

        private Appointment selectedAppointment;

        private TextBox txtId;
        private TextBox txtName;
        private TextBox txtPhone;
        private TextBox txtReason;
        private TextBox txtSearch;

        private DateTimePicker dtpAppointment;
        private DateTimePicker dtpFilterDate;

        private ComboBox cboDoctor;
        private ComboBox cboStatus;
        private ComboBox cboFilterStatus;

        private CheckBox chkFilterDate;

        private DataGridView dgvAppointments;
        private ToolStripStatusLabel statusLabel;

        public Form1()
        {
            InitializeComponent();
            CreateInterface();
            LoadSampleData();
            RefreshAppointmentList();
        }

        private void CreateInterface()
        {
            errorProvider.ContainerControl = this;

            MenuStrip menuStrip = new MenuStrip();

            ToolStripMenuItem menuFile =
                new ToolStripMenuItem("Hệ thống");

            ToolStripMenuItem menuExport =
                new ToolStripMenuItem("Xuất CSV");

            ToolStripMenuItem menuExit =
                new ToolStripMenuItem("Thoát");

            menuExport.Click += delegate
            {
                ExportCsv();
            };

            menuExit.Click += delegate
            {
                Close();
            };

            menuFile.DropDownItems.Add(menuExport);
            menuFile.DropDownItems.Add(menuExit);
            menuStrip.Items.Add(menuFile);

            StatusStrip statusStrip = new StatusStrip();

            statusLabel = new ToolStripStatusLabel();
            statusLabel.Text = "Sẵn sàng";

            statusStrip.Items.Add(statusLabel);

            TableLayoutPanel mainLayout =
                new TableLayoutPanel();

            mainLayout.Dock = DockStyle.Fill;
            mainLayout.Padding = new Padding(8);
            mainLayout.ColumnCount = 2;
            mainLayout.RowCount = 1;

            mainLayout.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 38F));

            mainLayout.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 62F));

            GroupBox inputGroup = new GroupBox();
            inputGroup.Text = "Thông tin lịch hẹn";
            inputGroup.Dock = DockStyle.Fill;
            inputGroup.Padding = new Padding(8);

            TableLayoutPanel inputLayout =
                new TableLayoutPanel();

            inputLayout.Dock = DockStyle.Fill;
            inputLayout.AutoScroll = true;
            inputLayout.ColumnCount = 2;
            inputLayout.RowCount = 0;
            inputLayout.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 35F));
            inputLayout.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 65F));

            txtId = new TextBox();
            txtName = new TextBox();
            txtPhone = new TextBox();
            txtReason = new TextBox();

            dtpAppointment = new DateTimePicker();
            dtpAppointment.Format = DateTimePickerFormat.Custom;
            dtpAppointment.CustomFormat = "dd/MM/yyyy HH:mm";
            dtpAppointment.ShowUpDown = true;
            dtpAppointment.Value = DateTime.Now.AddMinutes(30);

            cboDoctor = new ComboBox();
            cboDoctor.DropDownStyle = ComboBoxStyle.DropDownList;
            cboDoctor.Items.AddRange(new object[]
            {
                "BS. Nguyễn Văn A",
                "BS. Trần Thị B",
                "BS. Lê Văn C"
            });
            cboDoctor.SelectedIndex = 0;

            cboStatus = new ComboBox();
            cboStatus.DropDownStyle = ComboBoxStyle.DropDownList;
            cboStatus.Items.AddRange(new object[]
            {
                "Chờ khám",
                "Đang khám",
                "Hoàn tất",
                "Đã hủy"
            });
            cboStatus.SelectedIndex = 0;

            txtReason.Multiline = true;
            txtReason.ScrollBars = ScrollBars.Vertical;

            AddInputRow(
                inputLayout, "Mã lịch hẹn:", txtId, 38);

            AddInputRow(
                inputLayout, "Tên bệnh nhân:", txtName, 38);

            AddInputRow(
                inputLayout, "Số điện thoại:", txtPhone, 38);

            AddInputRow(
                inputLayout, "Ngày giờ khám:", dtpAppointment, 42);

            AddInputRow(
                inputLayout, "Bác sĩ:", cboDoctor, 38);

            AddInputRow(
                inputLayout, "Lý do khám:", txtReason, 78);

            AddInputRow(
                inputLayout, "Trạng thái:", cboStatus, 38);

            FlowLayoutPanel buttonPanel = new FlowLayoutPanel();
            buttonPanel.Dock = DockStyle.Fill;
            buttonPanel.AutoSize = false;
            buttonPanel.WrapContents = true;

            Button btnAdd = new Button();
            btnAdd.Text = "Thêm";
            btnAdd.Width = 75;
            btnAdd.Height = 32;

            Button btnUpdate = new Button();
            btnUpdate.Text = "Sửa";
            btnUpdate.Width = 75;
            btnUpdate.Height = 32;

            Button btnDelete = new Button();
            btnDelete.Text = "Xóa";
            btnDelete.Width = 75;
            btnDelete.Height = 32;

            Button btnClear = new Button();
            btnClear.Text = "Làm mới";
            btnClear.Width = 85;
            btnClear.Height = 32;

            buttonPanel.Controls.Add(btnAdd);
            buttonPanel.Controls.Add(btnUpdate);
            buttonPanel.Controls.Add(btnDelete);
            buttonPanel.Controls.Add(btnClear);

            int buttonRow = inputLayout.RowCount;
            inputLayout.RowCount++;
            inputLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 48));

            inputLayout.Controls.Add(buttonPanel, 0, buttonRow);
            inputLayout.SetColumnSpan(buttonPanel, 2);

            inputGroup.Controls.Add(inputLayout);

            TableLayoutPanel rightLayout =
                new TableLayoutPanel();

            rightLayout.Dock = DockStyle.Fill;
            rightLayout.ColumnCount = 1;
            rightLayout.RowCount = 3;

            rightLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 38));

            rightLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 42));

            rightLayout.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100F));

            FlowLayoutPanel searchPanel = new FlowLayoutPanel();
            searchPanel.Dock = DockStyle.Fill;
            searchPanel.WrapContents = false;
            searchPanel.AutoScroll = true;

            Label lblSearch = new Label();
            lblSearch.Text = "Tìm tên/SĐT:";
            lblSearch.Width = 90;
            lblSearch.TextAlign = ContentAlignment.MiddleLeft;
            lblSearch.Margin = new Padding(3, 5, 3, 3);

            txtSearch = new TextBox();
            txtSearch.Width = 250;
            txtSearch.Margin = new Padding(3, 5, 3, 3);

            searchPanel.Controls.Add(lblSearch);
            searchPanel.Controls.Add(txtSearch);

            FlowLayoutPanel filterPanel = new FlowLayoutPanel();
            filterPanel.Dock = DockStyle.Fill;
            filterPanel.WrapContents = false;
            filterPanel.AutoScroll = true;

            chkFilterDate = new CheckBox();
            chkFilterDate.Text = "Lọc ngày:";
            chkFilterDate.AutoSize = true;
            chkFilterDate.Margin = new Padding(3, 7, 3, 3);

            dtpFilterDate = new DateTimePicker();
            dtpFilterDate.Format = DateTimePickerFormat.Short;
            dtpFilterDate.Width = 115;
            dtpFilterDate.Enabled = false;
            dtpFilterDate.Margin = new Padding(3, 4, 3, 3);

            Label lblStatusFilter = new Label();
            lblStatusFilter.Text = "Trạng thái:";
            lblStatusFilter.AutoSize = true;
            lblStatusFilter.TextAlign = ContentAlignment.MiddleLeft;
            lblStatusFilter.Margin = new Padding(8, 7, 3, 3);

            cboFilterStatus = new ComboBox();
            cboFilterStatus.DropDownStyle =
                ComboBoxStyle.DropDownList;

            cboFilterStatus.Width = 130;
            cboFilterStatus.Items.AddRange(new object[]
            {
                "Tất cả",
                "Chờ khám",
                "Đang khám",
                "Hoàn tất",
                "Đã hủy"
            });

            cboFilterStatus.SelectedIndex = 0;
            cboFilterStatus.Margin = new Padding(3, 4, 3, 3);

            filterPanel.Controls.Add(chkFilterDate);
            filterPanel.Controls.Add(dtpFilterDate);
            filterPanel.Controls.Add(lblStatusFilter);
            filterPanel.Controls.Add(cboFilterStatus);

            dgvAppointments = new DataGridView();
            dgvAppointments.Dock = DockStyle.Fill;
            dgvAppointments.ReadOnly = true;
            dgvAppointments.AllowUserToAddRows = false;
            dgvAppointments.AllowUserToDeleteRows = false;
            dgvAppointments.AutoGenerateColumns = false;
            dgvAppointments.SelectionMode =
                DataGridViewSelectionMode.FullRowSelect;
            dgvAppointments.MultiSelect = false;
            dgvAppointments.RowHeadersVisible = false;
            dgvAppointments.AutoSizeColumnsMode =
                DataGridViewAutoSizeColumnsMode.Fill;
            dgvAppointments.BackgroundColor = Color.White;
            dgvAppointments.BorderStyle = BorderStyle.Fixed3D;

            AddGridColumn("Mã lịch", "AppointmentId", 80);
            AddGridColumn("Bệnh nhân", "PatientName", 120);
            AddGridColumn("Điện thoại", "Phone", 95);
            AddGridColumn("Ngày giờ khám", "AppointmentDate", 125,
                "dd/MM/yyyy HH:mm");
            AddGridColumn("Bác sĩ", "Doctor", 100);
            AddGridColumn("Trạng thái", "Status", 85);

            bindingSource.DataSource = displayedAppointments;
            dgvAppointments.DataSource = bindingSource;

            rightLayout.Controls.Add(searchPanel, 0, 0);
            rightLayout.Controls.Add(filterPanel, 0, 1);
            rightLayout.Controls.Add(dgvAppointments, 0, 2);

            mainLayout.Controls.Add(inputGroup, 0, 0);
            mainLayout.Controls.Add(rightLayout, 1, 0);

            this.Controls.Add(mainLayout);
            this.Controls.Add(statusStrip);
            this.Controls.Add(menuStrip);
            this.MainMenuStrip = menuStrip;

            btnAdd.Click += BtnAdd_Click;
            btnUpdate.Click += BtnUpdate_Click;
            btnDelete.Click += BtnDelete_Click;
            btnClear.Click += delegate
            {
                ClearInput();
            };

            txtSearch.TextChanged += delegate
            {
                RefreshAppointmentList();
            };

            chkFilterDate.CheckedChanged += delegate
            {
                dtpFilterDate.Enabled = chkFilterDate.Checked;
                RefreshAppointmentList();
            };

            dtpFilterDate.ValueChanged += delegate
            {
                RefreshAppointmentList();
            };

            cboFilterStatus.SelectedIndexChanged += delegate
            {
                RefreshAppointmentList();
            };

            dgvAppointments.CellClick += DgvAppointments_CellClick;
        }

        private void AddInputRow(
            TableLayoutPanel layout,
            string labelText,
            Control input,
            int rowHeight)
        {
            int row = layout.RowCount;
            layout.RowCount++;

            layout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, rowHeight));

            Label label = new Label();
            label.Text = labelText;
            label.Dock = DockStyle.Fill;
            label.TextAlign = ContentAlignment.MiddleLeft;
            label.Margin = new Padding(3);

            input.Dock = DockStyle.Fill;
            input.Margin = new Padding(3, 5, 3, 5);

            layout.Controls.Add(label, 0, row);
            layout.Controls.Add(input, 1, row);
        }

        private void AddGridColumn(
            string title,
            string propertyName,
            float fillWeight,
            string format = null)
        {
            DataGridViewTextBoxColumn column =
                new DataGridViewTextBoxColumn();

            column.HeaderText = title;
            column.DataPropertyName = propertyName;
            column.FillWeight = fillWeight;
            column.SortMode =
                DataGridViewColumnSortMode.Automatic;

            if (!string.IsNullOrEmpty(format))
            {
                column.DefaultCellStyle.Format = format;
            }

            dgvAppointments.Columns.Add(column);
        }

        private void LoadSampleData()
        {
            appointments.Add(new Appointment
            {
                AppointmentId = "LH001",
                PatientName = "Nguyễn Văn An",
                Phone = "0912345678",
                AppointmentDate = DateTime.Today.AddDays(1).AddHours(9),
                Doctor = "BS. Nguyễn Văn A",
                Reason = "Khám tổng quát",
                Status = "Chờ khám"
            });

            appointments.Add(new Appointment
            {
                AppointmentId = "LH002",
                PatientName = "Trần Thị Bình",
                Phone = "0987654321",
                AppointmentDate = DateTime.Today.AddDays(2).AddHours(14),
                Doctor = "BS. Trần Thị B",
                Reason = "Tái khám",
                Status = "Chờ khám"
            });
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ValidateInput())
            {
                MessageBox.Show(
                    "Vui lòng kiểm tra lại thông tin nhập.",
                    "Dữ liệu chưa hợp lệ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (FindAppointmentById(txtId.Text.Trim()) != null)
            {
                errorProvider.SetError(
                    txtId, "Mã lịch hẹn đã tồn tại.");

                MessageBox.Show(
                    "Mã lịch hẹn đã tồn tại. Hãy nhập mã khác.",
                    "Trùng mã",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            Appointment appointment = CreateAppointmentFromInput();
            appointments.Add(appointment);

            RefreshAppointmentList();
            ClearInput();

            MessageBox.Show(
                "Đã thêm lịch hẹn thành công.",
                "Thông báo",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void BtnUpdate_Click(object sender, EventArgs e)
        {
            if (selectedAppointment == null)
            {
                MessageBox.Show(
                    "Hãy chọn một lịch hẹn trong bảng trước khi sửa.",
                    "Chưa chọn lịch hẹn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (!ValidateInput())
            {
                MessageBox.Show(
                    "Vui lòng kiểm tra lại thông tin nhập.",
                    "Dữ liệu chưa hợp lệ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            string newId = txtId.Text.Trim();

            Appointment duplicate = FindAppointmentById(newId);

            if (duplicate != null &&
                !object.ReferenceEquals(duplicate, selectedAppointment))
            {
                errorProvider.SetError(
                    txtId, "Mã lịch hẹn đã được sử dụng.");

                MessageBox.Show(
                    "Mã lịch hẹn này thuộc về lịch hẹn khác.",
                    "Trùng mã",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            selectedAppointment.AppointmentId = newId;
            selectedAppointment.PatientName = txtName.Text.Trim();
            selectedAppointment.Phone = txtPhone.Text.Trim();
            selectedAppointment.AppointmentDate = dtpAppointment.Value;
            selectedAppointment.Doctor = cboDoctor.Text;
            selectedAppointment.Reason = txtReason.Text.Trim();
            selectedAppointment.Status = cboStatus.Text;

            RefreshAppointmentList();
            ClearInput();

            MessageBox.Show(
                "Đã cập nhật lịch hẹn thành công.",
                "Thông báo",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (selectedAppointment == null)
            {
                MessageBox.Show(
                    "Hãy chọn lịch hẹn cần xóa trong bảng.",
                    "Chưa chọn lịch hẹn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            DialogResult result = MessageBox.Show(
                "Bạn có chắc muốn xóa lịch hẹn của "
                + selectedAppointment.PatientName + "?",
                "Xác nhận xóa",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
            {
                return;
            }

            appointments.Remove(selectedAppointment);

            RefreshAppointmentList();
            ClearInput();

            MessageBox.Show(
                "Đã xóa lịch hẹn.",
                "Thông báo",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void DgvAppointments_CellClick(
            object sender,
            DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 ||
                e.RowIndex >= dgvAppointments.Rows.Count)
            {
                return;
            }

            Appointment appointment =
                dgvAppointments.Rows[e.RowIndex].DataBoundItem
                as Appointment;

            if (appointment == null)
            {
                return;
            }

            selectedAppointment = appointment;

            txtId.Text = appointment.AppointmentId;
            txtName.Text = appointment.PatientName;
            txtPhone.Text = appointment.Phone;
            dtpAppointment.Value = appointment.AppointmentDate;
            txtReason.Text = appointment.Reason;

            if (cboDoctor.Items.Contains(appointment.Doctor))
            {
                cboDoctor.SelectedItem = appointment.Doctor;
            }

            if (cboStatus.Items.Contains(appointment.Status))
            {
                cboStatus.SelectedItem = appointment.Status;
            }

            errorProvider.Clear();
        }

        private bool ValidateInput()
        {
            errorProvider.Clear();

            bool valid = true;

            if (string.IsNullOrWhiteSpace(txtId.Text))
            {
                errorProvider.SetError(
                    txtId, "Mã lịch hẹn không được để trống.");
                valid = false;
            }

            if (string.IsNullOrWhiteSpace(txtName.Text))
            {
                errorProvider.SetError(
                    txtName, "Tên bệnh nhân không được để trống.");
                valid = false;
            }

            string phone = txtPhone.Text.Trim();

            if (!Regex.IsMatch(phone, @"^0\d{9,10}$"))
            {
                errorProvider.SetError(
                    txtPhone,
                    "Số điện thoại phải bắt đầu bằng 0 và có 10 hoặc 11 chữ số.");
                valid = false;
            }

            if (dtpAppointment.Value < DateTime.Now)
            {
                errorProvider.SetError(
                    dtpAppointment,
                    "Ngày giờ khám không được nằm trong quá khứ.");
                valid = false;
            }

            if (cboDoctor.SelectedIndex < 0)
            {
                errorProvider.SetError(
                    cboDoctor, "Hãy chọn bác sĩ.");
                valid = false;
            }

            if (string.IsNullOrWhiteSpace(txtReason.Text))
            {
                errorProvider.SetError(
                    txtReason, "Vui lòng nhập lý do khám.");
                valid = false;
            }

            if (cboStatus.SelectedIndex < 0)
            {
                errorProvider.SetError(
                    cboStatus, "Hãy chọn trạng thái.");
                valid = false;
            }

            return valid;
        }

        private Appointment CreateAppointmentFromInput()
        {
            return new Appointment
            {
                AppointmentId = txtId.Text.Trim(),
                PatientName = txtName.Text.Trim(),
                Phone = txtPhone.Text.Trim(),
                AppointmentDate = dtpAppointment.Value,
                Doctor = cboDoctor.Text,
                Reason = txtReason.Text.Trim(),
                Status = cboStatus.Text
            };
        }

        private Appointment FindAppointmentById(string id)
        {
            foreach (Appointment appointment in appointments)
            {
                if (string.Equals(
                    appointment.AppointmentId,
                    id,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return appointment;
                }
            }

            return null;
        }

        private void RefreshAppointmentList()
        {
            displayedAppointments.RaiseListChangedEvents = false;
            displayedAppointments.Clear();

            string keyword = txtSearch.Text.Trim();
            string selectedStatus = cboFilterStatus.SelectedItem as string;

            foreach (Appointment appointment in appointments)
            {
                bool matchesKeyword =
                    string.IsNullOrEmpty(keyword) ||
                    appointment.PatientName.IndexOf(
                        keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    appointment.Phone.IndexOf(
                        keyword, StringComparison.OrdinalIgnoreCase) >= 0;

                bool matchesDate =
                    !chkFilterDate.Checked ||
                    appointment.AppointmentDate.Date == dtpFilterDate.Value.Date;

                bool matchesStatus =
                    string.IsNullOrEmpty(selectedStatus) ||
                    selectedStatus == "Tất cả" ||
                    appointment.Status == selectedStatus;

                if (matchesKeyword && matchesDate && matchesStatus)
                {
                    displayedAppointments.Add(appointment);
                }
            }

            displayedAppointments.RaiseListChangedEvents = true;
            bindingSource.ResetBindings(false);

            statusLabel.Text =
                "Số lịch hẹn đang hiển thị: "
                + displayedAppointments.Count
                + " / "
                + appointments.Count;
        }

        private void ClearInput()
        {
            selectedAppointment = null;

            txtId.Clear();
            txtName.Clear();
            txtPhone.Clear();
            txtReason.Clear();

            dtpAppointment.Value = DateTime.Now.AddMinutes(30);

            if (cboDoctor.Items.Count > 0)
            {
                cboDoctor.SelectedIndex = 0;
            }

            if (cboStatus.Items.Count > 0)
            {
                cboStatus.SelectedIndex = 0;
            }

            errorProvider.Clear();
            dgvAppointments.ClearSelection();
        }

        private void ExportCsv()
        {
            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Title = "Xuất danh sách lịch hẹn";
                dialog.Filter = "CSV (*.csv)|*.csv";
                dialog.FileName = "DanhSachLichHen.csv";

                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    using (StreamWriter writer = new StreamWriter(
                        dialog.FileName,
                        false,
                        new UTF8Encoding(true)))
                    {
                        writer.WriteLine(
                            "MaLich,TenBenhNhan,SoDienThoai,"
                            + "NgayGioKham,BacSi,LyDoKham,TrangThai");

                        foreach (Appointment appointment in displayedAppointments)
                        {
                            writer.WriteLine(
                                CsvEscape(appointment.AppointmentId) + "," +
                                CsvEscape(appointment.PatientName) + "," +
                                CsvEscape(appointment.Phone) + "," +
                                CsvEscape(appointment.AppointmentDate.ToString(
                                    "dd/MM/yyyy HH:mm")) + "," +
                                CsvEscape(appointment.Doctor) + "," +
                                CsvEscape(appointment.Reason) + "," +
                                CsvEscape(appointment.Status));
                        }
                    }

                    MessageBox.Show(
                        "Xuất file CSV thành công.",
                        "Thông báo",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Không thể xuất file.\n" + ex.Message,
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private string CsvEscape(string value)
        {
            if (value == null)
            {
                value = string.Empty;
            }

            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
    }
}
