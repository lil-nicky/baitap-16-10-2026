
using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace bai411
{
    public class RepairOrder
    {
        public string OrderId { get; set; }
        public string CustomerName { get; set; }
        public string Phone { get; set; }
        public string DeviceType { get; set; }
        public string ProblemDescription { get; set; }
        public DateTime ReceivedDate { get; set; }
        public decimal EstimatedCost { get; set; }
        public string Status { get; set; }
        public string ImagePath { get; set; }
    }

    public partial class Form1 : Form
    {
        private readonly BindingList<RepairOrder> orders =
            new BindingList<RepairOrder>();

        private readonly BindingList<RepairOrder> filteredOrders =
            new BindingList<RepairOrder>();

        private readonly BindingSource bindingSource =
            new BindingSource();

        private readonly ErrorProvider errorProvider =
            new ErrorProvider();

        private RepairOrder selectedOrder;
        private string selectedImagePath = null;

        private TextBox txtOrderId;
        private TextBox txtCustomerName;
        private TextBox txtPhone;
        private TextBox txtProblem;
        private TextBox txtSearch;

        private ComboBox cboDeviceType;
        private ComboBox cboStatus;
        private ComboBox cboFilterStatus;

        private DateTimePicker dtpReceivedDate;
        private NumericUpDown nudEstimatedCost;

        private PictureBox picDevice;
        private DataGridView dgvOrders;

        private ToolStripStatusLabel lblOrderCount;
        private ToolStripStatusLabel lblEstimatedTotal;

        public Form1()
        {
            InitializeComponent();
            CreateInterface();
            LoadSampleData();
            RefreshOrders();
        }

        // ==========================================
        // 1. TẠO GIAO DIỆN
        // ==========================================
        private void CreateInterface()
        {
            errorProvider.ContainerControl = this;

            // MenuStrip
            MenuStrip menuStrip = new MenuStrip();

            ToolStripMenuItem menuSystem =
                new ToolStripMenuItem("Hệ thống");

            ToolStripMenuItem menuExport =
                new ToolStripMenuItem("Xuất danh sách CSV");

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

            // StatusStrip
            StatusStrip statusStrip = new StatusStrip();

            lblOrderCount = new ToolStripStatusLabel();
            lblOrderCount.Text = "Số đơn đang hiển thị: 0";

            ToolStripStatusLabel separator =
                new ToolStripStatusLabel("|");

            lblEstimatedTotal = new ToolStripStatusLabel();
            lblEstimatedTotal.Text =
                "Tổng chi phí đơn chưa hoàn tất: 0";

            statusStrip.Items.Add(lblOrderCount);
            statusStrip.Items.Add(separator);
            statusStrip.Items.Add(lblEstimatedTotal);

            // Bố cục chính: 40% thông tin, 60% danh sách
            TableLayoutPanel mainLayout =
                new TableLayoutPanel();

            mainLayout.Dock = DockStyle.Fill;
            mainLayout.Padding = new Padding(8);
            mainLayout.ColumnCount = 2;
            mainLayout.RowCount = 1;

            mainLayout.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 40F));

            mainLayout.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 60F));

            // ==========================================
            // KHU VỰC NHẬP THÔNG TIN ĐƠN
            // ==========================================
            GroupBox inputGroup = new GroupBox();
            inputGroup.Text = "Thông tin đơn sửa chữa";
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

            txtOrderId = new TextBox();
            txtCustomerName = new TextBox();
            txtPhone = new TextBox();

            cboDeviceType = new ComboBox();
            cboDeviceType.DropDownStyle =
                ComboBoxStyle.DropDownList;

            cboDeviceType.Items.AddRange(new object[]
            {
                "Laptop",
                "Máy tính để bàn",
                "Điện thoại",
                "Máy tính bảng",
                "Máy in",
                "Màn hình",
                "Thiết bị khác"
            });

            cboDeviceType.SelectedIndex = 0;

            txtProblem = new TextBox();
            txtProblem.Multiline = true;
            txtProblem.ScrollBars = ScrollBars.Vertical;

            dtpReceivedDate = new DateTimePicker();
            dtpReceivedDate.Format = DateTimePickerFormat.Short;
            dtpReceivedDate.Value = DateTime.Today;

            nudEstimatedCost = new NumericUpDown();
            nudEstimatedCost.Minimum = 0;
            nudEstimatedCost.Maximum = 1000000000M;
            nudEstimatedCost.DecimalPlaces = 0;
            nudEstimatedCost.ThousandsSeparator = true;
            nudEstimatedCost.Increment = 100000;

            cboStatus = new ComboBox();
            cboStatus.DropDownStyle =
                ComboBoxStyle.DropDownList;

            cboStatus.Items.AddRange(new object[]
            {
                "Mới tiếp nhận",
                "Đang sửa",
                "Chờ linh kiện",
                "Hoàn tất",
                "Đã hủy"
            });

            cboStatus.SelectedIndex = 0;

            picDevice = new PictureBox();
            picDevice.Width = 130;
            picDevice.Height = 100;
            picDevice.BorderStyle = BorderStyle.FixedSingle;
            picDevice.SizeMode = PictureBoxSizeMode.Zoom;
            picDevice.BackColor = Color.White;

            Button btnChooseImage =
                CreateButton("Chọn ảnh", 90);

            FlowLayoutPanel imagePanel =
                new FlowLayoutPanel();

            imagePanel.Dock = DockStyle.Fill;
            imagePanel.WrapContents = false;
            imagePanel.AutoScroll = true;

            picDevice.Margin = new Padding(3, 3, 8, 3);
            btnChooseImage.Margin = new Padding(3, 30, 3, 3);

            imagePanel.Controls.Add(picDevice);
            imagePanel.Controls.Add(btnChooseImage);

            AddInputRow(
                inputLayout, "Mã đơn:", txtOrderId, 38);

            AddInputRow(
                inputLayout, "Tên khách hàng:", txtCustomerName, 38);

            AddInputRow(
                inputLayout, "Số điện thoại:", txtPhone, 38);

            AddInputRow(
                inputLayout, "Loại thiết bị:", cboDeviceType, 38);

            AddInputRow(
                inputLayout, "Mô tả lỗi:", txtProblem, 75);

            AddInputRow(
                inputLayout, "Ngày tiếp nhận:", dtpReceivedDate, 38);

            AddInputRow(
                inputLayout, "Chi phí dự kiến:", nudEstimatedCost, 38);

            AddInputRow(
                inputLayout, "Trạng thái:", cboStatus, 38);

            AddInputRow(
                inputLayout, "Ảnh thiết bị:", imagePanel, 110);

            // Các nút chức năng
            FlowLayoutPanel buttonPanel =
                new FlowLayoutPanel();

            buttonPanel.Dock = DockStyle.Fill;
            buttonPanel.WrapContents = true;
            buttonPanel.AutoScroll = true;

            Button btnAdd = CreateButton("Thêm", 65);
            Button btnUpdate = CreateButton("Cập nhật", 85);
            Button btnDelete = CreateButton("Xóa", 60);
            Button btnClear = CreateButton("Làm mới", 80);

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

            // ==========================================
            // KHU VỰC TÌM KIẾM VÀ BẢNG
            // ==========================================
            TableLayoutPanel rightLayout =
                new TableLayoutPanel();

            rightLayout.Dock = DockStyle.Fill;
            rightLayout.ColumnCount = 1;
            rightLayout.RowCount = 3;

            rightLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 42));

            rightLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, 42));

            rightLayout.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100F));

            // Tìm kiếm tức thời
            FlowLayoutPanel searchPanel =
                new FlowLayoutPanel();

            searchPanel.Dock = DockStyle.Fill;
            searchPanel.WrapContents = false;
            searchPanel.AutoScroll = true;

            Label lblSearch = new Label();
            lblSearch.Text = "Tìm mã/tên/SĐT:";
            lblSearch.AutoSize = true;
            lblSearch.Margin = new Padding(3, 8, 5, 3);

            txtSearch = new TextBox();
            txtSearch.Width = 230;
            txtSearch.Margin = new Padding(3, 5, 3, 3);

            searchPanel.Controls.Add(lblSearch);
            searchPanel.Controls.Add(txtSearch);

            // Lọc trạng thái
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

            cboFilterStatus.Width = 180;
            cboFilterStatus.Items.AddRange(new object[]
            {
                "Tất cả",
                "Mới tiếp nhận",
                "Đang sửa",
                "Chờ linh kiện",
                "Hoàn tất",
                "Đã hủy"
            });

            cboFilterStatus.SelectedIndex = 0;
            cboFilterStatus.Margin = new Padding(3, 4, 3, 3);

            filterPanel.Controls.Add(lblFilter);
            filterPanel.Controls.Add(cboFilterStatus);

            // DataGridView với cột cấu hình thủ công
            dgvOrders = new DataGridView();

            dgvOrders.Dock = DockStyle.Fill;
            dgvOrders.ReadOnly = true;
            dgvOrders.AllowUserToAddRows = false;
            dgvOrders.AllowUserToDeleteRows = false;
            dgvOrders.AutoGenerateColumns = false;
            dgvOrders.MultiSelect = false;
            dgvOrders.RowHeadersVisible = false;

            dgvOrders.SelectionMode =
                DataGridViewSelectionMode.FullRowSelect;

            dgvOrders.AutoSizeColumnsMode =
                DataGridViewAutoSizeColumnsMode.Fill;

            dgvOrders.BackgroundColor = Color.White;
            dgvOrders.BorderStyle = BorderStyle.Fixed3D;

            AddGridColumn("Mã đơn", "OrderId", 70);
            AddGridColumn("Khách hàng", "CustomerName", 100);
            AddGridColumn("Số điện thoại", "Phone", 90);
            AddGridColumn("Thiết bị", "DeviceType", 85);

            AddGridColumn(
                "Ngày tiếp nhận",
                "ReceivedDate",
                85,
                "dd/MM/yyyy");

            AddGridColumn(
                "Chi phí dự kiến",
                "EstimatedCost",
                100,
                "N0");

            AddGridColumn("Trạng thái", "Status", 100);
            AddGridColumn("Mô tả lỗi", "ProblemDescription", 130);

            bindingSource.DataSource = filteredOrders;
            dgvOrders.DataSource = bindingSource;

            rightLayout.Controls.Add(searchPanel, 0, 0);
            rightLayout.Controls.Add(filterPanel, 0, 1);
            rightLayout.Controls.Add(dgvOrders, 0, 2);

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
            btnChooseImage.Click += BtnChooseImage_Click;

            btnClear.Click += delegate
            {
                ClearInput();
            };

            txtSearch.TextChanged += delegate
            {
                RefreshOrders();
            };

            cboFilterStatus.SelectedIndexChanged += delegate
            {
                RefreshOrders();
            };

            dgvOrders.CellClick += DgvOrders_CellClick;
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
            string format = null)
        {
            DataGridViewTextBoxColumn column =
                new DataGridViewTextBoxColumn();

            column.HeaderText = header;
            column.DataPropertyName = propertyName;
            column.FillWeight = fillWeight;
            column.SortMode =
                DataGridViewColumnSortMode.Automatic;

            if (!string.IsNullOrEmpty(format))
            {
                column.DefaultCellStyle.Format = format;
            }

            dgvOrders.Columns.Add(column);
        }

        // ==========================================
        // 2. DỮ LIỆU MẪU
        // ==========================================
        private void LoadSampleData()
        {
            orders.Add(new RepairOrder
            {
                OrderId = "SC001",
                CustomerName = "Nguyễn Văn An",
                Phone = "0912345678",
                DeviceType = "Laptop",
                ProblemDescription = "Máy không lên nguồn",
                ReceivedDate = DateTime.Today.AddDays(-1),
                EstimatedCost = 850000M,
                Status = "Mới tiếp nhận",
                ImagePath = null
            });

            orders.Add(new RepairOrder
            {
                OrderId = "SC002",
                CustomerName = "Trần Thị Bình",
                Phone = "0987654321",
                DeviceType = "Điện thoại",
                ProblemDescription = "Màn hình bị vỡ",
                ReceivedDate = DateTime.Today,
                EstimatedCost = 1250000M,
                Status = "Đang sửa",
                ImagePath = null
            });

            orders.Add(new RepairOrder
            {
                OrderId = "SC003",
                CustomerName = "Lê Minh Cường",
                Phone = "0901234567",
                DeviceType = "Máy in",
                ProblemDescription = "Máy in không kéo giấy",
                ReceivedDate = DateTime.Today,
                EstimatedCost = 400000M,
                Status = "Hoàn tất",
                ImagePath = null
            });
        }

        // ==========================================
        // 3. CHỌN ẢNH THIẾT BỊ
        // ==========================================
        private void BtnChooseImage_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "Chọn ảnh thiết bị";
                dialog.Filter =
                    "Tệp ảnh|*.jpg;*.jpeg;*.png;*.bmp;*.gif";

                dialog.Multiselect = false;

                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    DisplayImage(dialog.FileName);
                    selectedImagePath = dialog.FileName;
                }
                catch (Exception ex)
                {
                    selectedImagePath = null;
                    ClearPicture();

                    MessageBox.Show(
                        "Không thể mở ảnh.\n" + ex.Message,
                        "Lỗi ảnh",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void DisplayImage(string path)
        {
            if (string.IsNullOrWhiteSpace(path) ||
                !File.Exists(path))
            {
                ClearPicture();
                return;
            }

            Image newImage;

            // Tạo bản sao để không khóa file ảnh gốc.
            using (Image source = Image.FromFile(path))
            {
                newImage = new Bitmap(source);
            }

            Image oldImage = picDevice.Image;
            picDevice.Image = newImage;
            picDevice.SizeMode = PictureBoxSizeMode.Zoom;

            if (oldImage != null)
            {
                oldImage.Dispose();
            }
        }

        private void ClearPicture()
        {
            Image oldImage = picDevice.Image;
            picDevice.Image = null;

            if (oldImage != null)
            {
                oldImage.Dispose();
            }
        }

        // ==========================================
        // 4. THÊM ĐƠN
        // ==========================================
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ValidateInput())
            {
                MessageBox.Show(
                    "Vui lòng kiểm tra các trường đang báo lỗi.",
                    "Dữ liệu không hợp lệ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (FindByOrderId(txtOrderId.Text.Trim()) != null)
            {
                errorProvider.SetError(
                    txtOrderId, "Mã đơn đã tồn tại.");

                MessageBox.Show(
                    "Mã đơn sửa chữa không được trùng.",
                    "Trùng mã đơn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            orders.Add(CreateOrderFromInput());

            RefreshOrders();
            ClearInput();

            MessageBox.Show(
                "Thêm đơn sửa chữa thành công.",
                "Thông báo",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        // ==========================================
        // 5. CẬP NHẬT ĐƠN
        // ==========================================
        private void BtnUpdate_Click(object sender, EventArgs e)
        {
            if (selectedOrder == null)
            {
                MessageBox.Show(
                    "Hãy chọn một đơn trong bảng trước khi cập nhật.",
                    "Chưa chọn đơn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (!ValidateInput())
            {
                MessageBox.Show(
                    "Vui lòng kiểm tra các trường đang báo lỗi.",
                    "Dữ liệu không hợp lệ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            RepairOrder duplicate =
                FindByOrderId(txtOrderId.Text.Trim());

            if (duplicate != null &&
                !object.ReferenceEquals(duplicate, selectedOrder))
            {
                errorProvider.SetError(
                    txtOrderId, "Mã đơn đã được sử dụng.");

                MessageBox.Show(
                    "Mã đơn này đã thuộc một đơn sửa chữa khác.",
                    "Trùng mã đơn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            selectedOrder.OrderId = txtOrderId.Text.Trim();
            selectedOrder.CustomerName = txtCustomerName.Text.Trim();
            selectedOrder.Phone = txtPhone.Text.Trim();
            selectedOrder.DeviceType = cboDeviceType.Text;
            selectedOrder.ProblemDescription = txtProblem.Text.Trim();
            selectedOrder.ReceivedDate = dtpReceivedDate.Value.Date;
            selectedOrder.EstimatedCost = nudEstimatedCost.Value;
            selectedOrder.Status = cboStatus.Text;
            selectedOrder.ImagePath = selectedImagePath;

            RefreshOrders();
            ClearInput();

            MessageBox.Show(
                "Cập nhật đơn sửa chữa thành công.",
                "Thông báo",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        // ==========================================
        // 6. XÓA ĐƠN CÓ XÁC NHẬN
        // ==========================================
        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (selectedOrder == null)
            {
                MessageBox.Show(
                    "Hãy chọn đơn cần xóa trong bảng.",
                    "Chưa chọn đơn",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            DialogResult result = MessageBox.Show(
                "Bạn có chắc muốn xóa đơn "
                + selectedOrder.OrderId + " của "
                + selectedOrder.CustomerName + "?",
                "Xác nhận xóa",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
            {
                return;
            }

            orders.Remove(selectedOrder);

            RefreshOrders();
            ClearInput();

            MessageBox.Show(
                "Đã xóa đơn sửa chữa.",
                "Thông báo",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        // ==========================================
        // 7. CHỌN DÒNG TRONG BẢNG
        // ==========================================
        private void DgvOrders_CellClick(
            object sender,
            DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 ||
                e.RowIndex >= dgvOrders.Rows.Count)
            {
                return;
            }

            RepairOrder order =
                dgvOrders.Rows[e.RowIndex].DataBoundItem
                as RepairOrder;

            if (order == null)
            {
                return;
            }

            selectedOrder = order;
            SetInputFromOrder(order);
        }

        private void SetInputFromOrder(RepairOrder order)
        {
            if (order == null)
            {
                return;
            }

            txtOrderId.Text = order.OrderId;
            txtCustomerName.Text = order.CustomerName;
            txtPhone.Text = order.Phone;
            txtProblem.Text = order.ProblemDescription;
            dtpReceivedDate.Value = order.ReceivedDate;

            nudEstimatedCost.Value = Math.Max(
                nudEstimatedCost.Minimum,
                Math.Min(nudEstimatedCost.Maximum, order.EstimatedCost));

            int deviceIndex =
                cboDeviceType.Items.IndexOf(order.DeviceType);

            if (deviceIndex >= 0)
            {
                cboDeviceType.SelectedIndex = deviceIndex;
            }
            else
            {
                cboDeviceType.SelectedIndex = -1;
            }

            int statusIndex =
                cboStatus.Items.IndexOf(order.Status);

            if (statusIndex >= 0)
            {
                cboStatus.SelectedIndex = statusIndex;
            }
            else
            {
                cboStatus.SelectedIndex = 0;
            }

            selectedImagePath = order.ImagePath;

            if (!string.IsNullOrWhiteSpace(selectedImagePath) &&
                File.Exists(selectedImagePath))
            {
                try
                {
                    DisplayImage(selectedImagePath);
                }
                catch
                {
                    ClearPicture();
                }
            }
            else
            {
                ClearPicture();
            }

            errorProvider.Clear();
        }

        // ==========================================
        // 8. KIỂM TRA DỮ LIỆU
        // ==========================================
        private bool ValidateInput()
        {
            errorProvider.Clear();

            bool valid = true;

            if (string.IsNullOrWhiteSpace(txtOrderId.Text))
            {
                errorProvider.SetError(
                    txtOrderId, "Mã đơn không được để trống.");
                valid = false;
            }

            if (string.IsNullOrWhiteSpace(txtCustomerName.Text))
            {
                errorProvider.SetError(
                    txtCustomerName,
                    "Tên khách hàng không được để trống.");
                valid = false;
            }

            if (string.IsNullOrWhiteSpace(txtPhone.Text))
            {
                errorProvider.SetError(
                    txtPhone,
                    "Số điện thoại không được để trống.");
                valid = false;
            }

            if (cboDeviceType.SelectedIndex < 0)
            {
                errorProvider.SetError(
                    cboDeviceType, "Hãy chọn loại thiết bị.");
                valid = false;
            }

            if (string.IsNullOrWhiteSpace(txtProblem.Text))
            {
                errorProvider.SetError(
                    txtProblem, "Mô tả lỗi là bắt buộc.");
                valid = false;
            }

            // NumericUpDown có Minimum = 0, không cho nhập giá trị âm.
            if (nudEstimatedCost.Value < 0)
            {
                errorProvider.SetError(
                    nudEstimatedCost,
                    "Chi phí dự kiến phải là số không âm.");
                valid = false;
            }

            if (cboStatus.SelectedIndex < 0)
            {
                errorProvider.SetError(
                    cboStatus, "Hãy chọn trạng thái đơn.");
                valid = false;
            }

            return valid;
        }

        private RepairOrder CreateOrderFromInput()
        {
            RepairOrder order = new RepairOrder();

            order.OrderId = txtOrderId.Text.Trim();
            order.CustomerName = txtCustomerName.Text.Trim();
            order.Phone = txtPhone.Text.Trim();
            order.DeviceType = cboDeviceType.Text;
            order.ProblemDescription = txtProblem.Text.Trim();
            order.ReceivedDate = dtpReceivedDate.Value.Date;
            order.EstimatedCost = nudEstimatedCost.Value;
            order.Status = cboStatus.Text;
            order.ImagePath = selectedImagePath;

            return order;
        }

        private RepairOrder FindByOrderId(string orderId)
        {
            foreach (RepairOrder order in orders)
            {
                if (string.Equals(
                    order.OrderId,
                    orderId,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return order;
                }
            }

            return null;
        }

        // ==========================================
        // 9. TÌM KIẾM, LỌC, THỐNG KÊ
        // ==========================================
        private void RefreshOrders()
        {
            string keyword = txtSearch.Text.Trim();
            string statusFilter =
                cboFilterStatus.SelectedItem as string;

            filteredOrders.RaiseListChangedEvents = false;
            filteredOrders.Clear();

            foreach (RepairOrder order in orders)
            {
                bool matchesKeyword =
                    string.IsNullOrEmpty(keyword) ||
                    ContainsText(order.OrderId, keyword) ||
                    ContainsText(order.CustomerName, keyword) ||
                    ContainsText(order.Phone, keyword);

                bool matchesStatus =
                    string.IsNullOrEmpty(statusFilter) ||
                    statusFilter == "Tất cả" ||
                    order.Status == statusFilter;

                if (matchesKeyword && matchesStatus)
                {
                    filteredOrders.Add(order);
                }
            }

            filteredOrders.RaiseListChangedEvents = true;
            bindingSource.ResetBindings(false);

            // Tổng chi phí của các đơn chưa hoàn tất
            decimal totalEstimatedCost = 0M;

            foreach (RepairOrder order in orders)
            {
                if (order.Status != "Hoàn tất")
                {
                    totalEstimatedCost += order.EstimatedCost;
                }
            }

            lblOrderCount.Text =
                "Số đơn đang hiển thị: " + filteredOrders.Count;

            lblEstimatedTotal.Text =
                "Tổng chi phí đơn chưa hoàn tất: "
                + totalEstimatedCost.ToString("N0");
        }

        private bool ContainsText(string source, string keyword)
        {
            if (source == null)
            {
                return false;
            }

            return source.IndexOf(
                keyword,
                StringComparison.OrdinalIgnoreCase) >= 0;
        }

        // ==========================================
        // 10. LÀM MỚI FORM
        // ==========================================
        private void ClearInput()
        {
            selectedOrder = null;
            selectedImagePath = null;

            txtOrderId.Clear();
            txtCustomerName.Clear();
            txtPhone.Clear();
            txtProblem.Clear();

            if (cboDeviceType.Items.Count > 0)
            {
                cboDeviceType.SelectedIndex = 0;
            }

            if (cboStatus.Items.Count > 0)
            {
                cboStatus.SelectedIndex = 0;
            }

            dtpReceivedDate.Value = DateTime.Today;
            nudEstimatedCost.Value = 0;

            ClearPicture();
            errorProvider.Clear();
            dgvOrders.ClearSelection();
        }

        // ==========================================
        // 11. XUẤT CSV
        // ==========================================
        private void ExportCsv()
        {
            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Title = "Xuất danh sách đơn sửa chữa";
                dialog.Filter = "CSV (*.csv)|*.csv";
                dialog.FileName = "DanhSachDonSuaChua.csv";
                dialog.OverwritePrompt = true;

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
                            "MaDon,TenKhachHang,SoDienThoai,LoaiThietBi,"
                            + "MoTaLoi,NgayTiepNhan,ChiPhiDuKien,TrangThai");

                        foreach (RepairOrder order in filteredOrders)
                        {
                            writer.WriteLine(
                                CsvEscape(order.OrderId) + "," +
                                CsvEscape(order.CustomerName) + "," +
                                CsvEscape(order.Phone) + "," +
                                CsvEscape(order.DeviceType) + "," +
                                CsvEscape(order.ProblemDescription) + "," +
                                CsvEscape(order.ReceivedDate.ToString(
                                    "dd/MM/yyyy")) + "," +
                                CsvEscape(order.EstimatedCost.ToString(
                                    "0")) + "," +
                                CsvEscape(order.Status));
                        }
                    }

                    MessageBox.Show(
                        "Đã xuất các đơn đang hiển thị ra file CSV.",
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
            {
                value = "";
            }

            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
    }
}
