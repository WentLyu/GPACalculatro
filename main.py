"""
该文件用于复旦大学数学科学学院本科生奖学金评定，作用是计算本科学年绩点折算得分，总分70分。
计算规则为，70*(总绩点+专业绩点)/8+修正分数，修正分数取决于学生是否在学年内修读了培养计划要求专业课程学分数的3/4以上，若修读足够学分，则修正分数=0，若否，则修正分数=-0.2。
Author: 吕文韬
Email: wtlv23@m.fudan.edu.cn
WeChat_ID: Went_Lyu
"""
import pandas as pd

# 要正常运行，只需要修改以下四个变量的值：
# dir_grades是包含所有同学成绩的Excel表所在目录
dir_grades = '/Users/admin/Desktop/scholarship/2021级全部成绩.xls'
# dir_nontrans表示专业课名单Excel表所在目录
dir_nontrans = '/Users/admin/Desktop/scholarship/数学学院专业成绩课程列表（计算专业绩点用-数学类）-20220908.xls'
# points表示该学年应该修读数学课程的分数
points = 40
# no_math_class_grade代表没有修读专业课程的同学专业课绩点赋值，此处设置为0
no_math_class_grade = 0


class Grades:
    def __init__(self, dir_grade, dir_nontran, gpoints) -> None:
        # 绩点-成绩换算表
        grade_point = {"A": 4.0,
                       "A-": 3.7,
                       "B+": 3.3,
                       "B": 3.0,
                       "B-": 2.7,
                       "C+": 2.3,
                       "C": 2.0,
                       "C-": 1.7,
                       "D": 1.3,
                       "D-": 1.0,
                       "F": 0.0,
                       "X": 0.0}

        # 输出挂科不能参评名单
        temp_df = pd.read_excel(dir_grade, index_col=None)
        f_man = temp_df[(temp_df['成绩'] == 'F') | (temp_df['成绩'] == 'NP') | (temp_df['成绩'] == 'X')].loc[:,
                ['学号', '姓名']].drop_duplicates()
        f_man.to_excel('./挂科不能参评名单.xlsx', index=False)

        # 删除拿了F的同学所在的行
        f_man = f_man['学号'].to_list()
        temp_df = temp_df[~temp_df['学号'].isin(f_man)]

        # 得到缓考成绩未出同学名单
        pending_man = temp_df[temp_df['成绩'] == '?'].loc[:, ['学号', '姓名']].drop_duplicates()
        pending_man.to_excel('./缓考成绩未出同学名单.xlsx', index=False)

        # 修改原文件，将所有PNP课程与缓考未出成绩课程去掉
        temp_df = temp_df[(temp_df['成绩'] != 'P') & (temp_df['成绩'] != '?')]
        temp_df['绩点'] = temp_df['成绩'].map(grade_point)
        temp_df['学分绩'] = temp_df['绩点'] * temp_df['学分']
        self.raw = temp_df
        self.dir_nontrans = dir_nontran
        self.point = gpoints

    @property
    def compute_fullgpa(self):
        temp = self.raw.groupby('学号')
        g = temp['学分绩'].sum()
        a = temp['学分'].sum()
        fullgpa = (g / a).to_frame()
        fullgpa.columns = ['总绩点']
        return fullgpa

    @property
    def compute_mathgpa(self):
        temp = self.raw
        courses = pd.read_excel(self.dir_nontrans, index_col=None)
        courses_code = courses['课程代码'].to_list()
        temp = temp[temp['课程代码'].isin(courses_code)].groupby('学号')
        g = temp['学分绩'].sum()
        a = temp['学分'].sum()
        math_gpa = (g / a).to_frame()

        # 该学年应修读数学专业课程学分数的0.75倍
        threshold = 0.75 * self.point

        # 计算修正分数
        a = a.apply(lambda x: -0.2 if x <= threshold else 0)

        math_gpa.columns = ['专业绩点']
        return [math_gpa, a]

    def calculate(self):
        fullgpa = self.compute_fullgpa
        [math_gpa, correction_point] = self.compute_mathgpa

        # 得到学号-姓名表
        id_name = self.raw.loc[:, ['学号', '姓名']].drop_duplicates()
        id_name = id_name.set_index('学号')

        # 补充没有修读专业课成绩的同学的专业成绩
        math_gpa = math_gpa.reindex(id_name.index, fill_value=no_math_class_grade)

        # 导出最后没有能够修读完指定专业课分数的同学名单
        correction_point = correction_point.reindex(id_name.index, fill_value=-0.2)
        id_name[correction_point != 0].to_excel('./未能修得足够专业课分数同学.xlsx')

        # 将专业绩点和总绩点合成一张表
        gpa_table = pd.concat([id_name, fullgpa, math_gpa], axis=1)

        # 计算最后的折算分数
        gpa_table['折算得分'] = (35 * (gpa_table['专业绩点'] + gpa_table['总绩点'])) / 4 + correction_point
        gpa_table.to_excel('./专业课折算得分.xlsx')


if __name__ == "__main__":
    # 以下代码进行计算
    grades = Grades(dir_grades, dir_nontrans, points)
    grades.calculate()
