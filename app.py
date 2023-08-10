from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd

app = Flask(__name__)

def process_excel(input_file, output_file, starting_number):
    # 读取Excel表格数据
    data = pd.read_excel(input_file)

    # 定义排序规则
    position_order = {'骨干': 1, '其他研究人员': 2}
    title_order = {'正高': 1, '副高': 2, '中级': 3, '初级': 4, '学生': 5}
    education_order = {'博士': 1, '硕士': 2, '本科': 3, '大专': 4}
    age_order = data['年龄'].max() - data['年龄']

    # 添加辅助列
    data['Position_Order'] = data['职务'].map(position_order)
    data['Title_Order'] = data['职称'].map(title_order)
    data['Education_Order'] = data['学历'].map(education_order)
    data['Age_Order'] = age_order

    # 根据排序规则进行排序
    sorted_data = data.sort_values(by=['Position_Order', 'Title_Order', 'Education_Order', 'Age_Order'],
                                   ascending=[True, True, True, False])

    # 移除辅助列
    sorted_data = sorted_data.drop(['Position_Order', 'Title_Order', 'Education_Order', 'Age_Order'], axis=1)

    # 添加排序序号
    sorted_data['排序序号'] = range(starting_number, starting_number + len(sorted_data))

    # 合并排序后的序号到原始数据
    result = pd.merge(data, sorted_data[['排序序号']], left_index=True, right_index=True, how='left')

    # 保存结果到Excel
    result.to_excel(output_file, index=False)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        input_file = request.files['input_file']
        starting_number = int(request.form['starting_number'])

        if input_file and input_file.filename.endswith('.xlsx'):
            input_filename = 'uploaded_file.xlsx'
            input_file.save(input_filename)

            output_filename = 'sorted_with_number.xlsx'
            process_excel(input_filename, output_filename, starting_number)

            return redirect(url_for('download', filename=output_filename))

    return render_template('index.html')

@app.route('/download/<filename>')
def download(filename):
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
