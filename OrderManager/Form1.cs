using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace OrderManager
{
    public partial class Form1 : Form
    {
        //数据库连接字符串
        private string conStr = "Data Source=DESKTOP-48LEQ9E\\SQLEXPRESS;Initial Catalog=OrderDB;Integrated Security=True";
        private SqlConnection con;
        private String[] customer = { "customerNo", "customerName", "telephone", "address", "zip" };
        private String[] employee = { "employeeNo", "employeeName", "gender", "birthday", "address", "telephone", "hireDate", "department", "headShip", "salary" };
        private String[] product = { "productNo", "productName", "productClass", "productPrice", "inStock" };
        private String[] ordermaster = { "orderNo", "customerNo", "employeeNo", "orderDate", "orderSum", "invoiceNo" };
        private String[] orderdetail = { "orderNo", "productNo", "quantity", "price" };
        private String[] all = { "","Customer.customerNo", "customerName", "Customer.telephone", "Customer.address", "zip", "Employee.employeeNo", "employeeName", "gender", "birthday", "Employee.address", "Employee.telephone", "hireDate", "department", "headShip", "salary", "Product.productNo", "productName", "productClass", "productPrice", "inStock", "OrderMaster.orderNo", "orderDate", "orderSum", "invoiceNo" };

        public Form1()
        {
            InitializeComponent();
            con = new SqlConnection(conStr);
            con.Open();
        }

        //处理字符串输入
        private String MT(String s)
        {
            return s != "" ? "'" + s + "'" : "NULL";
        }

        //释放资源
        private void DataRelease(object sender, FormClosedEventArgs e)
        {
            con.Close();
        }

        //查询客户信息
        private void button1_Click(object sender, EventArgs e)
        {
            String comStr = "select * from Customer";
            if(textBox1.Text!=""||textBox2.Text!=""||textBox3.Text!=""||textBox4.Text!=""||textBox5.Text!="")
            {
                comStr += " where";
                if (textBox1.Text != "")
                    comStr += " " + customer[0] + " = '" + textBox1.Text + "' and";
                if (textBox2.Text != "")
                    comStr += " " + customer[1] + " = '" + textBox2.Text + "' and";
                if (textBox3.Text != "")
                    comStr += " " + customer[2] + " = '" + textBox3.Text + "' and";
                if (textBox4.Text != "")
                    comStr += " " + customer[3] + " = '" + textBox4.Text + "' and";
                if (textBox5.Text != "")
                    comStr += " " + customer[4] + " = '" + textBox5.Text + "' and";
                comStr = comStr.Substring(0, comStr.Length - 4);
            }
            SqlDataAdapter da = new SqlDataAdapter(comStr, con);
            DataTable dt = new DataTable();
            da.Fill(dt);  //查询结果记录在数据表中
            this.dataGridView1.DataSource = dt;
            this.dataGridView1.Columns[0].ReadOnly = true;
            addButtonCol(this.dataGridView1);
            this.dataGridView1.Update();
        }
        //添加两列button，分别对应更新和删除操作
        private void addButtonCol(DataGridView e)
        {
            for (int i = 0; i < e.Columns.Count; i++ )
                if (e.Columns[i] is DataGridViewButtonColumn)
                {
                    e.Columns.RemoveAt(i);
                    e.Columns.RemoveAt(i);
                    break;
                }
            DataGridViewButtonColumn button_update_Col = new DataGridViewButtonColumn();
            button_update_Col.Name = "ButtonUpdate";
            button_update_Col.Text = "更新";
            button_update_Col.ReadOnly = true;
            button_update_Col.Width = 80;
            e.Columns.Add(button_update_Col);
            for (int i = 0; i < e.Rows.Count; i++)
                e.Rows[i].Cells["ButtonUpdate"].Value = "更新";
            DataGridViewButtonColumn button_del_Col = new DataGridViewButtonColumn();
            button_del_Col.Name = "ButtonDel";
            button_del_Col.Text = "删除";
            button_del_Col.ReadOnly = true;
            button_del_Col.Width = 80;
            e.Columns.Add(button_del_Col);
            for (int i = 0; i < e.Rows.Count; i++)
                e.Rows[i].Cells["ButtonDel"].Value = "删除";
        }

        //修改、删除客户信息
        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (dataGridView1.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex > -1)
                {
                    int row = e.RowIndex;
                    if (this.dataGridView1.CurrentCell.Value.ToString() == "更新")
                    {                        
                        String sql = "update Customer set ";
                        for (int i = 1; i < customer.Length; i++)
                            sql += customer[i] + " = '" + this.dataGridView1[i, row].Value + "', "; 
                        sql = sql.Substring(0, sql.Length - 2);
                        sql += " where customerNo = '" + this.dataGridView1[0, row].Value + "'";
                        SqlCommand com = new SqlCommand(sql, con);
                        try
                        {
                            com.ExecuteNonQuery();
                            MessageBox.Show("更新成功");
                        }
                        catch
                        {
                            MessageBox.Show("输入信息格式错误，更新失败");
                        }
                    }
                    else
                    {
                        String sql = "delete from Customer where customerNo = '" + this.dataGridView1[0, row].Value + "'";
                        SqlCommand com = new SqlCommand(sql, con);
                        try
                        {
                            com.ExecuteNonQuery();
                            MessageBox.Show("删除成功");
                        }
                        catch
                        {
                            MessageBox.Show("与该客户还有订单关系，故不能删除该客户");
                        }
                    }
                }
            }
            catch 
            {
                MessageBox.Show("操作错误");
            }
        }

        //查询员工信息
        private void button2_Click(object sender, EventArgs e)
        {
            String comStr = "select * from Employee";
            if (textBox6.Text != "" || textBox7.Text != "" || textBox8.Text != "" || textBox9.Text != "" || textBox10.Text != ""
                || textBox11.Text != "" || textBox12.Text != "" || textBox13.Text != "" || textBox14.Text != "" || textBox15.Text != "")
            {
                comStr += " where";
                if (textBox6.Text != "")
                    comStr += " " + employee[0] + " = '" + textBox6.Text + "' and";
                if (textBox7.Text != "")
                    comStr += " " + employee[1] + " = '" + textBox7.Text + "' and";
                if (textBox8.Text != "")
                    comStr += " " + employee[2] + " = '" + textBox8.Text + "' and";
                if (textBox9.Text != "")
                    comStr += " " + employee[3] + " = '" + textBox9.Text + "' and";
                if (textBox10.Text != "")
                    comStr += " " + employee[4] + " = '" + textBox10.Text + "' and";
                if (textBox11.Text != "")
                    comStr += " " + employee[5] + " = '" + textBox11.Text + "' and";
                if (textBox12.Text != "")
                    comStr += " " + employee[6] + " = '" + textBox12.Text + "' and";
                if (textBox13.Text != "")
                    comStr += " " + employee[7] + " = '" + textBox13.Text + "' and";
                if (textBox14.Text != "")
                    comStr += " " + employee[8] + " = '" + textBox14.Text + "' and";
                if (textBox15.Text != "")
                    comStr += " " + employee[9] + " = '" + textBox15.Text + "' and";
                comStr = comStr.Substring(0, comStr.Length - 4);
            }
            SqlDataAdapter da = new SqlDataAdapter(comStr, con);
            DataTable dt = new DataTable();
            try {
                da.Fill(dt);
                this.dataGridView2.DataSource = dt;
                this.dataGridView2.Columns[0].ReadOnly = true;
                addButtonCol(this.dataGridView2);
                this.dataGridView2.Update();
            }
            catch
            {
                MessageBox.Show("格式错误");
            }
        }

        //修改、删除员工信息
        private void DataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (dataGridView2.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex > -1)
                {
                    int row = e.RowIndex;
                    if (this.dataGridView2.CurrentCell.Value.ToString() == "更新")
                    {
                        String sql = "update Employee set ";
                        for (int i = 1; i < employee.Length; i++)
                            sql += employee[i] + " = '" + this.dataGridView2[i, row].Value + "', ";
                        sql = sql.Substring(0, sql.Length - 2);
                        sql += " where employeeNo = '" + this.dataGridView2[0, row].Value + "'";
                        SqlCommand com = new SqlCommand(sql, con);
                        try
                        {
                            com.ExecuteNonQuery();
                            MessageBox.Show("更新成功");
                        }
                        catch
                        {
                            MessageBox.Show("输入信息格式错误，更新失败");
                        }
                    }
                    else
                    {
                        String sql = "delete from Employee where employeeNo = '" + this.dataGridView2[0, row].Value + "'";
                        SqlCommand com = new SqlCommand(sql, con);
                        try
                        {
                            com.ExecuteNonQuery();
                            MessageBox.Show("删除成功");
                        }
                        catch
                        {
                            MessageBox.Show("与该员工还有订单关系，故不能删除该员工");
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("操作错误");
            }
        }

        //查询产品信息
        private void button3_Click(object sender, EventArgs e)
        {
            String comStr = "select * from Product";
            if (textBox16.Text != "" || textBox17.Text != "" || textBox18.Text != "" || textBox19.Text != "" || textBox20.Text != "")
            {
                comStr += " where";
                if (textBox16.Text != "")
                    comStr += " " + product[0] + " = '" + textBox16.Text + "' and";
                if (textBox17.Text != "")
                    comStr += " " + product[1] + " = '" + textBox17.Text + "' and";
                if (textBox18.Text != "")
                    comStr += " " + product[2] + " = '" + textBox18.Text + "' and";
                if (textBox19.Text != "")
                    comStr += " " + product[3] + " = '" + textBox19.Text + "' and";
                if (textBox20.Text != "")
                    comStr += " " + product[4] + " = '" + textBox20.Text + "' and";
                comStr = comStr.Substring(0, comStr.Length - 4);
            }
            SqlDataAdapter da = new SqlDataAdapter(comStr, con);
            DataTable dt = new DataTable();
            try {
                da.Fill(dt);
                this.dataGridView3.DataSource = dt;
                this.dataGridView3.Columns[0].ReadOnly = true;
                addButtonCol(this.dataGridView3);
                this.dataGridView3.Update();
            }
            catch
            {
                MessageBox.Show("格式错误");
            }
        }

        //修改、删除员工信息
        private void DataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (dataGridView3.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex > -1)
                {
                    int row = e.RowIndex;
                    if (this.dataGridView3.CurrentCell.Value.ToString() == "更新")
                    {
                        String sql = "update Product set ";
                        for (int i = 1; i < product.Length; i++)
                            sql += product[i] + " = '" + this.dataGridView3[i, row].Value + "', ";
                        sql = sql.Substring(0, sql.Length - 2);
                        sql += " where productNo = '" + this.dataGridView3[0, row].Value + "'";
                        SqlCommand com = new SqlCommand(sql, con);
                        try
                        {
                            com.ExecuteNonQuery();
                            MessageBox.Show("更新成功");
                        }
                        catch
                        {
                            MessageBox.Show("输入信息格式错误，更新失败");
                        }
                    }
                    else
                    {
                        String sql = "delete from Product where productNo = '" + this.dataGridView3[0, row].Value + "'";
                        SqlCommand com = new SqlCommand(sql, con);
                        try
                        {
                            com.ExecuteNonQuery();
                            MessageBox.Show("删除成功");
                        }
                        catch
                        {
                            MessageBox.Show("与该产品还有订单关系，故不能删除该产品");
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("操作错误");
            }
        }

        //查询订单信息
        private void button4_Click(object sender, EventArgs e)
        {
            String comStr = "select * from OrderMaster left join OrderDetail on OrderMaster.orderNo = OrderDetail.orderNo";
            if (textBox21.Text != "" || textBox22.Text != "" || textBox23.Text != "" || textBox24.Text != "" || textBox25.Text != "")
            {
                comStr += " where";
                if (textBox21.Text != "")
                    comStr += " " + ordermaster[0] + " = '" + textBox21.Text + "' and";
                if (textBox22.Text != "")
                    comStr += " " + ordermaster[1] + " = '" + textBox22.Text + "' and";
                if (textBox23.Text != "")
                    comStr += " " + ordermaster[2] + " = '" + textBox23.Text + "' and";
                if (textBox24.Text != "")
                    comStr += " " + ordermaster[3] + " = '" + textBox24.Text + "' and";
                if (textBox25.Text != "")
                    comStr += " " + ordermaster[4] + " = '" + textBox25.Text + "' and";
                comStr = comStr.Substring(0, comStr.Length - 4);
            }
            SqlDataAdapter da = new SqlDataAdapter(comStr, con);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                this.dataGridView4.DataSource = dt;
                this.dataGridView4.ReadOnly = true;
                this.dataGridView4.Update();
            }
            catch
            {
                MessageBox.Show("格式错误");
            }
        }

        //新增客户
        private void button5_Click(object sender, EventArgs e)
        {
            if(textBox1.Text==""||textBox2.Text==""||textBox3.Text==""||textBox4.Text=="")
            {
                MessageBox.Show("客户信息不全");
            }
            else
            {
                String comStr = "insert into Customer values (" + MT(textBox1.Text) + "," + MT(textBox2.Text) + "," + MT(textBox3.Text) + "," + MT(textBox4.Text) + "," + MT(textBox5.Text) + ")";
                SqlCommand com = new SqlCommand(comStr, con);
                try
                {
                    if (com.ExecuteNonQuery() > 0)
                        MessageBox.Show("新增成功");
                }
                catch
                {
                    MessageBox.Show("客户信息格式错误，或已存在该客户");
                }
            }
        }

        //新增员工
        private void button6_Click(object sender, EventArgs e)
        {
            if (textBox6.Text == "")
            {
                MessageBox.Show("有空值");
            }
            else
            {
                String comStr = "insert into Employee values (" + MT(textBox6.Text) + "," + MT(textBox7.Text) + "," + MT(textBox8.Text) + "," + MT(textBox9.Text) + "," + MT(textBox10.Text) + "," +MT(textBox11.Text) + "," + MT(textBox12.Text) + "," + MT(textBox13.Text) + "," + MT(textBox14.Text) + "," + MT(textBox15.Text) + ")";
                SqlCommand com = new SqlCommand(comStr, con);
                try
                {
                    if (com.ExecuteNonQuery() > 0)
                        MessageBox.Show("新增成功");
                }
                catch
                {
                    MessageBox.Show("格式错误");
                }
            }
        }

        //新增产品
        private void button7_Click(object sender, EventArgs e)
        {
            if (textBox16.Text == ""|| textBox17.Text == "" || textBox18.Text == "" || textBox19.Text == "")
            {
                MessageBox.Show("有空值");
            }
            else
            {
                String comStr = "insert into Product values (" + MT(textBox16.Text) + "," + MT(textBox17.Text) + "," + MT(textBox18.Text) + "," + MT(textBox19.Text) + "," + MT(textBox20.Text) + ")";
                SqlCommand com = new SqlCommand(comStr, con);
                try
                {
                    if (com.ExecuteNonQuery() > 0)
                        MessageBox.Show("新增成功");
                }
                catch
                {
                    MessageBox.Show("格式错误");
                }
            }
        }

        //级联删除订单
        private void OrderDelete(object sender, EventArgs e)
        {
            String orderNo = dataGridView4[0, dataGridView4.SelectedCells[0].RowIndex].Value + "";
            String comStr1 = "delete from OrderDetail where OrderNo = '" + orderNo + "'";
            String comStr2 = "delete from OrderMaster where OrderNo = '" + orderNo + "'";
            SqlTransaction st = con.BeginTransaction();  //开启事务
            SqlCommand com = new SqlCommand(comStr1, con, st);
            try {
                com.ExecuteNonQuery();
                com.CommandText = comStr2;
                com.ExecuteNonQuery();
                st.Commit();
                MessageBox.Show("删除成功");
            }
            catch
            {
                st.Rollback();
                MessageBox.Show("删除失败");
            }
            finally
            {
                st.Dispose();
            }
        }

        //组合查询
        private void button8_Click(object sender, EventArgs e)
        {
            //要查询的内容
            String comStr = "select distinct ";
            if (checkedListBox1.GetItemChecked(0))
                comStr += "Customer.*,";
            if (checkedListBox1.GetItemChecked(1))
                comStr += "Employee.*,";
            if (checkedListBox1.GetItemChecked(2))
                comStr += "Product.*,";
            if (checkedListBox1.GetItemChecked(3))
                comStr += "OrderMaster.*,";
            if (checkedListBox1.GetItemChecked(4))
                comStr += "OrderDetail.*,";
            if(comStr.EndsWith(" "))
            {
                MessageBox.Show("请选择查询内容");
                return;
            }
            comStr = comStr.Substring(0, comStr.Length - 1);

            //联接表
            comStr += " from Customer,Employee,Product,OrderMaster,OrderDetail"
                + " where OrderMaster.customerNo = Customer.customerNo"
                + " and OrderMaster.employeeNo = Employee.employeeNo"
                + " and OrderMaster.OrderNo = OrderDetail.OrderNo"
                + " and OrderDetail.ProductNo = Product.ProductNo";

            //查询条件
            if(comboBox1.SelectedIndex > 0)
            {
                comStr += " and " + all[comboBox1.SelectedIndex] + " = '" + textBox27.Text + "'";
            }
            if (comboBox2.SelectedIndex > 0)
            {
                comStr += " and " + all[comboBox2.SelectedIndex] + " = '" + textBox28.Text + "'";
            }
            if (comboBox3.SelectedIndex > 0)
            {
                comStr += " and " + all[comboBox3.SelectedIndex] + " = '" + textBox29.Text + "'";
            }
            if (comboBox4.SelectedIndex > 0)
            {
                comStr += " and " + all[comboBox4.SelectedIndex] + " = '" + textBox30.Text + "'";
            }

            //查询
            SqlDataAdapter da = new SqlDataAdapter(comStr, con);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                this.dataGridView5.DataSource = dt;
                this.dataGridView5.Update();
            }
            catch
            {
                MessageBox.Show("格式错误");
            }
        }

        //新增订单
        private void button9_Click(object sender, EventArgs e)
        {
            if(textBox31.Text==""&& textBox32.Text == ""&& textBox33.Text == "")
            {
                MessageBox.Show("请输入商品编号");
                return;
            }

            //查询订单所含商品信息
            int c = 0;
            decimal cost = 0;
            String comStr = "select productNo,productPrice,inStock from Product where productNo in (";
            if (textBox31.Text != "")
            {
                c++;
                comStr += "'" + textBox31.Text + "',";
            }
            if (textBox32.Text != "")
            {
                c++;
                comStr += "'" + textBox32.Text + "',";
            }
            if (textBox33.Text != "")
            {
                c++;
                comStr += "'" + textBox33.Text + "',";
            }
            comStr = comStr.Substring(0, comStr.Length - 1);
            comStr += ")";
            SqlDataAdapter da = new SqlDataAdapter(comStr, con);
            DataTable dt = new DataTable();
            da.Fill(dt);

            //根据查询结果处理不正确的商品编号及库存不足情况
            //并更改商品剩余情况计算订单总额
            if(dt.Rows.Count!=c)
            {
                MessageBox.Show("不正确的商品编号");
                return;
            }
            for(int i=0;i<c;i++)
            {
                if(dt.Rows[i][0]+""==textBox31.Text)
                {
                    if((int)dt.Rows[i][2]<numericUpDown1.Value)
                    {
                        MessageBox.Show("商品库存不足");
                        return;
                    }
                    cost += numericUpDown1.Value * (decimal)dt.Rows[i][1];
                }
                else if (dt.Rows[i][0] + "" == textBox32.Text)
                {
                    if ((int)dt.Rows[i][2] < numericUpDown2.Value)
                    {
                        MessageBox.Show("商品库存不足");
                        return;
                    }
                    cost += numericUpDown2.Value * (decimal)dt.Rows[i][1];
                }
                else if (dt.Rows[i][0] + "" == textBox33.Text)
                {
                    if ((int)dt.Rows[i][2] < numericUpDown3.Value)
                    {
                        MessageBox.Show("商品库存不足");
                        return;
                    }
                    cost += numericUpDown3.Value * (decimal)dt.Rows[i][1];
                }
            }

            //开启事务
            SqlTransaction st = con.BeginTransaction();
            SqlCommand com = new SqlCommand("", con, st);
            String orderNo;

            //建立主表
            while (true)
            {
                //时间戳作为orderNo
                orderNo = DateTime.Now.ToString("yyyyMMddmmss");
                String comStr1 = "insert into OrderMaster (OrderNo) values ('" + orderNo + "')";
                com.CommandText = comStr1;
                try
                {
                    com.ExecuteNonQuery();
                    break;
                }
                catch
                {
                    //orderNo重复，则继续循环，直至取得合适的时间戳
                }
            }

            //修改主表其他信息
            try {
                if (textBox34.Text != "")
                {
                    com.CommandText = "update OrderMaster set CustomerNo = '" + textBox34.Text + "' where OrderNo = '" + orderNo + "'";
                    com.ExecuteNonQuery();
                }
                if (textBox35.Text != "")
                {
                    com.CommandText = "update OrderMaster set EmployeeNo = '" + textBox35.Text + "' where OrderNo = '" + orderNo + "'";
                    com.ExecuteNonQuery();
                }
                com.CommandText = "update OrderMaster set orderDate = getDate() where OrderNo = '" + orderNo + "'";
                com.ExecuteNonQuery();
                com.CommandText = "update OrderMaster set orderSum = " + cost + " where OrderNo = '" + orderNo + "'";
                com.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("不正确的客户编号或业务员编号");
                st.Rollback();
                return;
            }

            //修改副表信息和商品库存
            try
            {
                for (int i = 0; i < c; i++)
                {
                    if (dt.Rows[i][0] + "" == textBox31.Text)
                    {
                        com.CommandText = "insert into OrderDetail values ('" + orderNo + "','" + dt.Rows[i][0] + "'," + (int)numericUpDown1.Value + "," + dt.Rows[i][1] + ")";
                        com.ExecuteNonQuery();
                        com.CommandText = "update Product set inStock = inStock - " + (int)numericUpDown1.Value + " where ProductNo = '" + textBox31.Text + "'";
                        com.ExecuteNonQuery();
                    }
                    else if (dt.Rows[i][0] + "" == textBox32.Text)
                    {
                        com.CommandText = "insert into OrderDetail values ('" + orderNo + "','" + dt.Rows[i][0] + "'," + (int)numericUpDown2.Value + "," + dt.Rows[i][1] + ")";
                        com.ExecuteNonQuery();
                        com.CommandText = "update Product set inStock = inStock - " + (int)numericUpDown2.Value + " where ProductNo = '" + textBox32.Text + "'";
                        com.ExecuteNonQuery();
                    }
                    else if (dt.Rows[i][0] + "" == textBox33.Text)
                    {
                        com.CommandText = "insert into OrderDetail values ('" + orderNo + "','" + dt.Rows[i][0] + "'," + (int)numericUpDown3.Value + "," + dt.Rows[i][1] + ")";
                        com.ExecuteNonQuery();
                        com.CommandText = "update Product set inStock = inStock - " + (int)numericUpDown3.Value + " where ProductNo = '" + textBox33.Text + "'";
                        com.ExecuteNonQuery();
                    }
                }
            }
            catch
            {
                st.Rollback();
                return;
            }

            st.Commit();
            MessageBox.Show("新建订单成功");
        }

    }
}
