
using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace bai410
{
    public class BorrowRecord
    {
        public string BorrowId { get; set; }
        public string StudentId { get; set; }
        public string StudentName { get; set; }
        public string Equipment { get; set; }
        public int Quantity { get; set; }
        public DateTime BorrowDate { get; set; }
        public DateTime DueDate { get; set; }
        public DateTime? ActualReturnDate { get; set; }

        public string Status
        {
            get
            {
                if (ActualReturnDate.HasValue)
                    return "Đã trả";

                if (DueDate.Date < DateTime.Today)
                    return "Quá hạn";

                return "Đang mượn";
            }
        }
    }

    public partial class Form1 : Form
    {
        private readonly BindingList<BorrowRecord> records =
            new BindingList<BorrowRecord>();

        private readonly BindingList<BorrowRecord> filteredRecords =
            new BindingList<BorrowRecord>();

        private readonly BindingSource bindingSource =
            new BindingSource();

        private readonly ErrorProvider errorProvider =
            new ErrorProvider();

        private BorrowRecord selectedRecord;

        private TextBox txtBorrowId;
        private TextBox txtStudentId;
        private TextBox txtStudentName;
        private TextBox txtSearch;

        private ComboBox cboEquipment;
        private ComboBox cboFilterStatus;

        private NumericUpDown nudQuantity;

        private DateTimePicker dtpBorrowDate;
        private DateTimePicker dtpDueDate;
        private DateTimePicker dtpReturnDate;

        private CheckBox chkReturned;

        private DataGridView dgvRecords;

        private ToolStripStatusLabel lblTotal;
        private ToolStripStatusLabel lblOverdue;

        public Form1()
        {
            InitializeComponent();
            CreateInterface();
            LoadSampleData();
            RefreshRecords();
        }

        // =========================
        // 1. TẠO GIAO DIỆN
        // =========================
        private void CreateInterface()
        {
            errorProvider.ContainerControl = this;

            MenuStrip menuStrip = new MenuStrip();

            ToolStripMenuItem menuSystem =
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

            menuSystem.DropDownItems.Add(menuExport);
            menuSystem.DropDownItems.Add(menuExit);
            menuStrip.Items.Add(menuSystem);

            StatusStrip statusStrip = new StatusStrip();

            lblTotal = new ToolStripStatusLabel();
            lblTotal.Text = "Tổng số phiếu: 0";

            ToolStripStatusLabel separator =
                new ToolStripStatusLabel("|");

            lblOverdue = new ToolStripStatusLabel();
            lblOverdue.Text = "Số phiếu quá hạn: 0";

            statusStrip.Items.Add(lblTotal);
            statusStrip.Items.Add(separator);
            statusStrip.Items.Add(lblOverdue);

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

            // Khu vực nhập dữ liệu
            GroupBox inputGroup = new GroupBox();
            inputGroup.Text = "Thông tin phiếu mượn";
            inputGroup.Dock = DockStyle.Fill;
            inputGroup.Padding = new Padding(8);

            TableLayoutPanel inputLayout =
                new TableLayoutPanel();

            inputLayout.Dock = DockStyle.Fill;
            inputLayout.AutoScroll = true;
            inputLayout.ColumnCount = 2;
            inputLayout.RowCount = 0;

            inputLayout.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 36F));

            inputLayout.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 64F));

            txtBorrowId = new TextBox();
            txtStudentId = new TextBox();
            txtStudentName = new TextBox();

            cboEquipment = new ComboBox();
            cboEquipment.DropDownStyle =
                ComboBoxStyle.DropDownList;

            cboEquipment.Items.AddRange(new object[]
            {
                "Laptop Dell",
                "Arduino Uno",
                "Raspberry Pi",
                "Máy chiếu",
                "Kit cảm biến",
                "Đồng hồ vạn năng",
                "Bộ thí nghiệm điện"
            });

            if (cboEquipment.Items.Count > 0)
                cboEquipment.SelectedIndex = 0;

            nudQuantity = new NumericUpDown();
            nudQuantity.Minimum = 1;
            nudQuantity.Maximum = 10000;
            nudQuantity.Value = 1;
            nudQuantity.ThousandsSeparator = true;

            dtpBorrowDate = new DateTimePicker();
            dtpBorrowDate.Format = DateTimePickerFormat.Short;
            dtpBorrowDate.Value = DateTime.Today;

            dtpDueDate = new DateTimePicker();
            dtpDueDate.Format = DateTimePickerFormat.Short;
            dtpDueDate.Value = DateTime.Today.AddDays(7);

            dtpReturnDate = new DateTimePicker();
            dtpReturnDate.Format = DateTimePickerFormat.Short;
            dtpReturnDate.Value = DateTime.Today;
            dtpReturnDate.Enabled = false;

            FlowLayoutPanel returnPanel =
                new FlowLayoutPanel();

            returnPanel.Dock = DockStyle.Fill;
            returnPanel.WrapContents = false;
            returnPanel.AutoScroll = true;

            chkReturned = new CheckBox();
            chkReturned.Text = "Đã trả";
            chkReturned.AutoSize = true;
            chkReturned.Margin = new Padding(3, 7, 5, 3);

            dtpReturnDate.Width = 125;
            dtpReturnDate.Margin = new Padding(3, 4, 3, 3);

            returnPanel.Controls.Add(chkReturned);
            returnPanel.Controls.Add(dtpReturnDate);

            AddInputRow(
                inputLayout, "Mã phiếu:", txtBorrowId, 40);

            AddInputRow(
                inputLayout, "Mã sinh viên:", txtStudentId, 40);

            AddInputRow(
                inputLayout, "Họ tên sinh viên:", txtStudentName, 40);

            AddInputRow(
                inputLayout, "Thiết bị:", cboEquipment, 40);

            AddInputRow(
                inputLayout, "Số lượng:", nudQuantity, 40);

            AddInputRow(
                inputLayout, "Ngày mượn:", dtpBorrowDate, 40);

            AddInputRow(
                inputLayout, "Hạn trả:", dtpDueDate, 40);

            AddInputRow(
                inputLayout, "Ngày trả thực tế:", returnPanel, 45);

            FlowLayoutPanel buttonPanel =
                new FlowLayoutPanel();

            buttonPanel.Dock = DockStyle.Fill;
            buttonPanel.WrapContents = true;
            buttonPanel.AutoScroll = true;

            Button btnAdd = CreateButton("Thêm", 65);
            Button btnUpdate = CreateButton("Cập nhật", 80);
            Button btnDelete = CreateButton("Xóa", 60);
            Button btnClear = CreateButton("Làm mới", 80);
            Button btnMarkReturned =
                CreateButton("Đánh dấu đã trả", 125);

            buttonPanel.Controls.Add(btnAdd);
            buttonPanel.Controls.Add(btnUpdate);
            buttonPanel.Controls.Add(btnDelete);
            buttonPanel.Controls.Add(btnClear);
            buttonPanel.Controls.Add(btnMarkReturned);

            int buttonRow = inputLayout.RowCount;

            inputLayout.RowCount++;
            inputLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 90));

            inputLayout.Controls.Add(buttonPanel, 0, buttonRow);
            inputLayout.SetColumnSpan(buttonPanel, 2);

            inputGroup.Controls.Add(inputLayout);

            // Khu vực tìm kiếm, lọc và bảng
            TableLayoutPanel rightLayout =
                new TableLayoutPanel();

            rightLayout.Dock = DockStyle.Fill;
            rightLayout.ColumnCount = 1;
            rightLayout.RowCount = 3;

            rightLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 40));

            rightLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 42));

            rightLayout.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100F));

            FlowLayoutPanel searchPanel =
                new FlowLayoutPanel();

            searchPanel.Dock = DockStyle.Fill;
            searchPanel.WrapContents = false;
            searchPanel.AutoScroll = true;

            Label lblSearch = new Label();
            lblSearch.Text = "Tìm mã SV/thiết bị:";
            lblSearch.AutoSize = true;
            lblSearch.Margin = new Padding(3, 8, 5, 3);

            txtSearch = new TextBox();
            txtSearch.Width = 220;
            txtSearch.Margin = new Padding(3, 5, 5, 3);

            Button btnExport = CreateButton("Xuất CSV", 85);
            btnExport.Height = 28;
            btnExport.Margin = new Padding(3, 3, 3, 3);

            searchPanel.Controls.Add(lblSearch);
            searchPanel.Controls.Add(txtSearch);
            searchPanel.Controls.Add(btnExport);

            FlowLayoutPanel filterPanel =
                new FlowLayoutPanel();

            filterPanel.Dock = DockStyle.Fill;
            filterPanel.WrapContents = false;
            filterPanel.AutoScroll = true;

            Label lblFilter = new Label();
            lblFilter.Text = "Lọc trạng thái:";
            lblFilter.AutoSize = true;
            lblFilter.Margin = new Padding(3, 8, 5, 3);

            cboFilterStatus = new ComboBox();
            cboFilterStatus.DropDownStyle =
                ComboBoxStyle.DropDownList;

            cboFilterStatus.Width = 160;
            cboFilterStatus.Items.AddRange(new object[]
            {
                "Tất cả",
                "Đang mượn",
                "Quá hạn",
                "Đã trả"
            });

            cboFilterStatus.SelectedIndex = 0;
            cboFilterStatus.Margin = new Padding(3, 4, 3, 3);

            filterPanel.Controls.Add(lblFilter);
            filterPanel.Controls.Add(cboFilterStatus);

            dgvRecords = new DataGridView();
            dgvRecords.Dock = DockStyle.Fill;
            dgvRecords.ReadOnly = true;
            dgvRecords.AllowUserToAddRows = false;
            dgvRecords.AllowUserToDeleteRows = false;
            dgvRecords.AutoGenerateColumns = false;
            dgvRecords.MultiSelect = false;
            dgvRecords.RowHeadersVisible = false;

            dgvRecords.SelectionMode =
                DataGridViewSelectionMode.FullRowSelect;

            dgvRecords.AutoSizeColumnsMode =
                DataGridViewAutoSizeColumnsMode.Fill;

            dgvRecords.BackgroundColor = Color.White;
            dgvRecords.BorderStyle = BorderStyle.Fixed3D;

            // Cấu hình cột thủ công
            AddGridColumn("Mã phiếu", "BorrowId", 75);
            AddGridColumn("Mã SV", "StudentId", 75);
            AddGridColumn("Họ tên sinh viên", "StudentName", 115);
            AddGridColumn("Thiết bị", "Equipment", 105);
            AddGridColumn("SL", "Quantity", 40);
            AddGridColumn(
                "Ngày mượn", "BorrowDate", 85, "dd/MM/yyyy");
            AddGridColumn(
                "Hạn trả", "DueDate", 85, "dd/MM/yyyy");
            AddGridColumn(
                "Ngày trả thực tế",
                "ActualReturnDate",
                95,
                "dd/MM/yyyy",
                "Chưa trả");
            AddGridColumn("Tình trạng", "Status", 75);

            bindingSource.DataSource = filteredRecords;
            dgvRecords.DataSource = bindingSource;

            rightLayout.Controls.Add(searchPanel, 0, 0);
            rightLayout.Controls.Add(filterPanel, 0, 1);
            rightLayout.Controls.Add(dgvRecords, 0, 2);

            mainLayout.Controls.Add(inputGroup, 0, 0);
            mainLayout.Controls.Add(rightLayout, 1, 0);

            this.Controls.Add(mainLayout);
            this.Controls.Add(statusStrip);
            this.Controls.Add(menuStrip);
            this.MainMenuStrip = menuStrip;

            // Đăng ký sự kiện
            btnAdd.Click += BtnAdd_Click;
            btnUpdate.Click += BtnUpdate_Click;
            btnDelete.Click += BtnDelete_Click;
            btnMarkReturned.Click += BtnMarkReturned_Click;

            btnClear.Click += delegate
            {
                ClearInput();
            };

            btnExport.Click += delegate
            {
                ExportCsv();
            };

            txtSearch.TextChanged += delegate
            {
                RefreshRecords();
            };

            cboFilterStatus.SelectedIndexChanged += delegate
            {
                RefreshRecords();
            };

            chkReturned.CheckedChanged += delegate
            {
                dtpReturnDate.Enabled = chkReturned.Checked;
            };

            dgvRecords.CellClick += DgvRecords_CellClick;
        }

        private Button CreateButton(string text, int width)
        {
            Button button = new Button();
            button.Text = text;
            button.Width = width;
            button.Height = 32;
            button.Margin = new Padding(3, 5, 3, 3);
            return button;
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
            string header,
            string propertyName,
            float fillWeight,
            string format = null,
            string nullValue = null)
        {
            DataGridViewTextBoxColumn column =
                new DataGridViewTextBoxColumn();

            column.HeaderText = header;
            column.DataPropertyName = propertyName;
            column.FillWeight = fillWeight;
            column.SortMode =
                DataGridViewColumnSortMode.Automatic;

            if (!string.IsNullOrEmpty(format))
                column.DefaultCellStyle.Format = format;

            if (nullValue != null)
                column.DefaultCellStyle.NullValue = nullValue;

            dgvRecords.Columns.Add(column);
        }

        // =========================
        // 2. DỮ LIỆU MẪU
        // =========================
        private void LoadSampleData()
        {
            records.Add(new BorrowRecord
            {
                BorrowId = "PM001",
                StudentId = "SV001",
                StudentName = "Nguyễn Văn An",
                Equipment = "Laptop Dell",
                Quantity = 1,
                BorrowDate = DateTime.Today.AddDays(-2),
                DueDate = DateTime.Today.AddDays(5),
                ActualReturnDate = null
            });

            records.Add(new BorrowRecord
            {
                BorrowId = "PM002",
                StudentId = "SV002",
                StudentName = "Trần Thị Bình",
                Equipment = "Arduino Uno",
                Quantity = 2,
                BorrowDate = DateTime.Today.AddDays(-8),
                DueDate = DateTime.Today.AddDays(-2),
                ActualReturnDate = null
            });

            records.Add(new BorrowRecord
            {
                BorrowId = "PM003",
                StudentId = "SV003",
                StudentName = "Lê Minh Cường",
                Equipment = "Kit cảm biến",
                Quantity = 1,
                BorrowDate = DateTime.Today.AddDays(-6),
                DueDate = DateTime.Today.AddDays(-1),
                ActualReturnDate = DateTime.Today
            });
        }

        // =========================
        // 3. THÊM PHIẾU
        // =========================
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ValidateInput())
            {
                MessageBox.Show(
                    "Hãy kiểm tra những trường đang báo lỗi.",
                    "Dữ liệu không hợp lệ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (FindByBorrowId(txtBorrowId.Text.Trim()) != null)
            {
                errorProvider.SetError(
                    txtBorrowId, "Mã phiếu đã tồn tại.");

                MessageBox.Show(
                    "Mã phiếu mượn không được trùng.",
                    "Trùng mã phiếu",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            records.Add(CreateRecordFromInput());

            RefreshRecords();
            ClearInput();

            MessageBox.Show(
                "Thêm phiếu mượn thành công.",
                "Thông báo",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        // =========================
        // 4. CẬP NHẬT PHIẾU
        // =========================
        private void BtnUpdate_Click(object sender, EventArgs e)
        {
            if (selectedRecord == null)
            {
                MessageBox.Show(
                    "Hãy chọn một phiếu trong bảng trước khi cập nhật.",
                    "Chưa chọn phiếu",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (!ValidateInput())
            {
                MessageBox.Show(
                    "Hãy kiểm tra những trường đang báo lỗi.",
                    "Dữ liệu không hợp lệ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            BorrowRecord duplicate =
                FindByBorrowId(txtBorrowId.Text.Trim());

            if (duplicate != null &&
                !object.ReferenceEquals(duplicate, selectedRecord))
            {
                errorProvider.SetError(
                    txtBorrowId, "Mã phiếu đã được sử dụng.");

                MessageBox.Show(
                    "Mã phiếu này thuộc một phiếu khác.",
                    "Trùng mã phiếu",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            selectedRecord.BorrowId = txtBorrowId.Text.Trim();
            selectedRecord.StudentId = txtStudentId.Text.Trim();
            selectedRecord.StudentName = txtStudentName.Text.Trim();
            selectedRecord.Equipment = cboEquipment.Text;
            selectedRecord.Quantity = (int)nudQuantity.Value;
            selectedRecord.BorrowDate = dtpBorrowDate.Value.Date;
            selectedRecord.DueDate = dtpDueDate.Value.Date;

            if (chkReturned.Checked)
            {
                selectedRecord.ActualReturnDate =
                    dtpReturnDate.Value.Date;
            }
            else
            {
                selectedRecord.ActualReturnDate = null;
            }

            RefreshRecords();
            ClearInput();

            MessageBox.Show(
                "Cập nhật phiếu mượn thành công.",
                "Thông báo",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        // =========================
        // 5. XÓA PHIẾU
        // =========================
        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (selectedRecord == null)
            {
                MessageBox.Show(
                    "Hãy chọn phiếu cần xóa trong bảng.",
                    "Chưa chọn phiếu",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            DialogResult result = MessageBox.Show(
                "Bạn có chắc muốn xóa phiếu "
                + selectedRecord.BorrowId + "?",
                "Xác nhận xóa",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
                return;

            records.Remove(selectedRecord);

            RefreshRecords();
            ClearInput();

            MessageBox.Show(
                "Đã xóa phiếu mượn.",
                "Thông báo",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        // =========================
        // 6. ĐÁNH DẤU ĐÃ TRẢ
        // =========================
        private void BtnMarkReturned_Click(object sender, EventArgs e)
        {
            if (selectedRecord == null)
            {
                MessageBox.Show(
                    "Hãy chọn phiếu mượn cần đánh dấu đã trả.",
                    "Chưa chọn phiếu",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (selectedRecord.ActualReturnDate.HasValue)
            {
                MessageBox.Show(
                    "Phiếu này đã được trả trước đó.",
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            DateTime returnDate = DateTime.Today;

            if (returnDate < selectedRecord.BorrowDate.Date)
            {
                errorProvider.SetError(
                    dtpReturnDate,
                    "Ngày trả không được trước ngày mượn.");

                MessageBox.Show(
                    "Ngày hiện tại trước ngày mượn nên không thể đánh dấu đã trả.",
                    "Ngày trả không hợp lệ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            selectedRecord.ActualReturnDate = returnDate;

            // Bỏ bộ lọc trạng thái để người dùng nhìn thấy phiếu đã trả.
            cboFilterStatus.SelectedIndex = 0;

            RefreshRecords();
            SetInputFromRecord(selectedRecord);
            SelectRecordInGrid(selectedRecord);

            MessageBox.Show(
                "Đã cập nhật ngày trả thực tế và trạng thái thành Đã trả.",
                "Thông báo",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        // =========================
        // 7. CHỌN DÒNG TRONG BẢNG
        // =========================
        private void DgvRecords_CellClick(
            object sender,
            DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 ||
                e.RowIndex >= dgvRecords.Rows.Count)
                return;

            BorrowRecord record =
                dgvRecords.Rows[e.RowIndex].DataBoundItem
                as BorrowRecord;

            if (record == null)
                return;

            selectedRecord = record;
            SetInputFromRecord(record);
        }

        private void SetInputFromRecord(BorrowRecord record)
        {
            if (record == null)
                return;

            txtBorrowId.Text = record.BorrowId;
            txtStudentId.Text = record.StudentId;
            txtStudentName.Text = record.StudentName;

            int equipmentIndex =
                cboEquipment.Items.IndexOf(record.Equipment);

            if (equipmentIndex >= 0)
                cboEquipment.SelectedIndex = equipmentIndex;
            else
                cboEquipment.SelectedIndex = -1;

            nudQuantity.Value = Math.Max(
                nudQuantity.Minimum,
                Math.Min(nudQuantity.Maximum, record.Quantity));

            dtpBorrowDate.Value = record.BorrowDate;
            dtpDueDate.Value = record.DueDate;

            if (record.ActualReturnDate.HasValue)
            {
                chkReturned.Checked = true;
                dtpReturnDate.Value =
                    record.ActualReturnDate.Value;
            }
            else
            {
                chkReturned.Checked = false;
                dtpReturnDate.Value = DateTime.Today;
            }

            errorProvider.Clear();
        }

        private void SelectRecordInGrid(BorrowRecord record)
        {
            dgvRecords.ClearSelection();

            for (int i = 0; i < dgvRecords.Rows.Count; i++)
            {
                BorrowRecord rowRecord =
                    dgvRecords.Rows[i].DataBoundItem as BorrowRecord;

                if (object.ReferenceEquals(rowRecord, record))
                {
                    dgvRecords.Rows[i].Selected = true;

                    if (dgvRecords.Rows[i].Cells.Count > 0)
                    {
                        dgvRecords.CurrentCell =
                            dgvRecords.Rows[i].Cells[0];
                    }

                    break;
                }
            }
        }

        // =========================
        // 8. KIỂM TRA DỮ LIỆU
        // =========================
        private bool ValidateInput()
        {
            errorProvider.Clear();

            bool valid = true;

            if (string.IsNullOrWhiteSpace(txtBorrowId.Text))
            {
                errorProvider.SetError(
                    txtBorrowId, "Mã phiếu không được để trống.");
                valid = false;
            }

            if (string.IsNullOrWhiteSpace(txtStudentId.Text))
            {
                errorProvider.SetError(
                    txtStudentId, "Mã sinh viên không được để trống.");
                valid = false;
            }

            if (string.IsNullOrWhiteSpace(txtStudentName.Text))
            {
                errorProvider.SetError(
                    txtStudentName, "Họ tên sinh viên không được để trống.");
                valid = false;
            }

            if (cboEquipment.SelectedIndex < 0 ||
                string.IsNullOrWhiteSpace(cboEquipment.Text))
            {
                errorProvider.SetError(
                    cboEquipment, "Hãy chọn thiết bị.");
                valid = false;
            }

            if (nudQuantity.Value <= 0)
            {
                errorProvider.SetError(
                    nudQuantity, "Số lượng phải lớn hơn 0.");
                valid = false;
            }

            if (dtpDueDate.Value.Date < dtpBorrowDate.Value.Date)
            {
                errorProvider.SetError(
                    dtpDueDate,
                    "Hạn trả không được trước ngày mượn.");
                valid = false;
            }

            if (chkReturned.Checked &&
                dtpReturnDate.Value.Date < dtpBorrowDate.Value.Date)
            {
                errorProvider.SetError(
                    dtpReturnDate,
                    "Ngày trả thực tế không được trước ngày mượn.");
                valid = false;
            }

            return valid;
        }

        private BorrowRecord CreateRecordFromInput()
        {
            BorrowRecord record = new BorrowRecord();

            record.BorrowId = txtBorrowId.Text.Trim();
            record.StudentId = txtStudentId.Text.Trim();
            record.StudentName = txtStudentName.Text.Trim();
            record.Equipment = cboEquipment.Text;
            record.Quantity = (int)nudQuantity.Value;
            record.BorrowDate = dtpBorrowDate.Value.Date;
            record.DueDate = dtpDueDate.Value.Date;

            if (chkReturned.Checked)
                record.ActualReturnDate = dtpReturnDate.Value.Date;
            else
                record.ActualReturnDate = null;

            return record;
        }

        private BorrowRecord FindByBorrowId(string borrowId)
        {
            foreach (BorrowRecord record in records)
            {
                if (string.Equals(
                    record.BorrowId,
                    borrowId,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return record;
                }
            }

            return null;
        }

        // =========================
        // 9. TÌM KIẾM, LỌC, THỐNG KÊ
        // =========================
        private void RefreshRecords()
        {
            string keyword = txtSearch.Text.Trim();
            string statusFilter =
                cboFilterStatus.SelectedItem as string;

            filteredRecords.RaiseListChangedEvents = false;
            filteredRecords.Clear();

            foreach (BorrowRecord record in records)
            {
                bool matchesKeyword =
                    string.IsNullOrEmpty(keyword) ||
                    record.StudentId.IndexOf(
                        keyword,
                        StringComparison.OrdinalIgnoreCase) >= 0 ||
                    record.Equipment.IndexOf(
                        keyword,
                        StringComparison.OrdinalIgnoreCase) >= 0;

                bool matchesStatus =
                    string.IsNullOrEmpty(statusFilter) ||
                    statusFilter == "Tất cả" ||
                    record.Status == statusFilter;

                if (matchesKeyword && matchesStatus)
                    filteredRecords.Add(record);
            }

            filteredRecords.RaiseListChangedEvents = true;
            bindingSource.ResetBindings(false);

            int overdueCount = 0;

            foreach (BorrowRecord record in records)
            {
                if (record.Status == "Quá hạn")
                    overdueCount++;
            }

            lblTotal.Text = "Tổng số phiếu: " + records.Count;
            lblOverdue.Text = "Số phiếu quá hạn: " + overdueCount;
        }

        // =========================
        // 10. LÀM MỚI
        // =========================
        private void ClearInput()
        {
            selectedRecord = null;

            txtBorrowId.Clear();
            txtStudentId.Clear();
            txtStudentName.Clear();

            if (cboEquipment.Items.Count > 0)
                cboEquipment.SelectedIndex = 0;

            nudQuantity.Value = 1;

            dtpBorrowDate.Value = DateTime.Today;
            dtpDueDate.Value = DateTime.Today.AddDays(7);
            dtpReturnDate.Value = DateTime.Today;

            chkReturned.Checked = false;
            dtpReturnDate.Enabled = false;

            errorProvider.Clear();
            dgvRecords.ClearSelection();
        }

        // =========================
        // 11. XUẤT CSV
        // =========================
        private void ExportCsv()
        {
            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Title = "Xuất danh sách phiếu mượn";
                dialog.Filter = "CSV (*.csv)|*.csv";
                dialog.FileName = "DanhSachPhieuMuon.csv";
                dialog.OverwritePrompt = true;

                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                try
                {
                    using (StreamWriter writer = new StreamWriter(
                        dialog.FileName,
                        false,
                        new UTF8Encoding(true)))
                    {
                        writer.WriteLine(
                            "MaPhieu,MaSinhVien,HoTenSinhVien,ThietBi,"
                            + "SoLuong,NgayMuon,HanTra,NgayTraThucTe,TinhTrang");

                        foreach (BorrowRecord record in filteredRecords)
                        {
                            string returnDate = "";

                            if (record.ActualReturnDate.HasValue)
                            {
                                returnDate =
                                    record.ActualReturnDate.Value.ToString(
                                        "dd/MM/yyyy");
                            }

                            writer.WriteLine(
                                CsvEscape(record.BorrowId) + "," +
                                CsvEscape(record.StudentId) + "," +
                                CsvEscape(record.StudentName) + "," +
                                CsvEscape(record.Equipment) + "," +
                                record.Quantity + "," +
                                CsvEscape(record.BorrowDate.ToString(
                                    "dd/MM/yyyy")) + "," +
                                CsvEscape(record.DueDate.ToString(
                                    "dd/MM/yyyy")) + "," +
                                CsvEscape(returnDate) + "," +
                                CsvEscape(record.Status));
                        }
                    }

                    MessageBox.Show(
                        "Đã xuất các phiếu đang được lọc ra file CSV.",
                        "Xuất CSV thành công",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Không thể xuất file CSV.\n" + ex.Message,
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private string CsvEscape(string value)
        {
            if (value == null)
                value = "";

            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
    }
}
