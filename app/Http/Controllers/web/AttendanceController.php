<?php

namespace App\Http\Controllers\web;

use Illuminate\Http\Request;

use App\Http\Requests;
use App\Http\Controllers\Controller;

class AttendanceController extends Controller
{
    /**
     * @param Request $request
     * @return string
     * excel output
     * @author atshike <atshike@foxmail.com>
     */
    public function excel(Request $request)
    {
        $mothday = static::getDaysAndWeek(date('Y-m', strtotime($request->month)));
        $daily = 'Daily' . date('YmdHis');
        $rows['test'] = 'test';
        \Maatwebsite\Excel\Facades\Excel::create($daily, function ($excel) use ($request, $mothday) {
            $excel->sheet('Daily', function ($excel) use ($request, $mothday) {
                $excel->mergeCells('A1:O4')->cell('A1', function ($cells) use ($request) {
                    static::setCell($cells, date('Y年m月', strtotime($request->get('month'))) . '出勤记录');
                    $cells->setFontSize(16);
                });
                $i = 0;
                while ($i < count($mothday)) {
                    $y = $i + 7;
                    static::setMergeNo($excel, $y);
                    $excel->row($y, array(
                        //数据填充
                        $mothday[$i], '',   //时间
                        '', '',             //分类
                        '', '',             //开始
                        '', '',             //结束
                        '', '',             //休息
                        '',                 //平日
                        '',                 //休息
                        '',                 //加班
                        '',                 //规定时间
                        '',                 //迟到早退
                        '',                 //终差
                        '',                 //扣除
                        '', '', '',         //扣除理由
                        ''                  //备注
                    ))->setHeight(array(
                        $y => 14,
                    ));
                    static::setCellWidth($excel);
                    $i++;
                }
                static::setMergeCells($excel, 'E' . ($y + 4) . ':H' . ($y + 4), 'E' . ($y + 4), '加班【】');
                static::overTimeHours($excel, $y);
                static::excelNo($excel, $y);
                $excel->setBorder('A1:Y' . $y, 'thin');
                $excel->setBorder('A' . ($y + 4) . ':C' . ($y + 6), 'thin');
                $excel->setBorder('W' . ($y + 3) . ':Y' . ($y + 6), 'thin');
                $excel->setBorder('E' . ($y + 4) . ':H' . ($y + 6), 'thin');
            });
        })->download('xlsx');
    }

    /**
     * 遍历月天数
     * @author atshike <atshike@foxmail.com>
     */
    public static function getDaysAndWeek($date)
    {
        $year = date('Y', strtotime($date));
        $month = date('m', strtotime($date));
        $firstday = date('Y-m-01', strtotime($date));
        $daynum = date('t', strtotime($firstday));
        $everydays = array();
        for ($i = 1; $i <= $daynum; $i++) {
            $everydays[] = $year . '-' . $month . '-' . static::dayzero($i);
        }
        return $everydays;
    }

    /**
     * getDaysAndWeek
     * @author atshike <atshike@foxmail.com>
     */
    public static function dayzero($result)
    {
        if ($result <= 9) {
            return '0' . $result;
        }
        return $result;
    }

    /**
     * Set Font Size
     * excel
     * @author atshike <atshike@foxmail.com>
     */
    private function setCell($cells, $value)
    {
        $cells->setValue($value);
        $cells->setFontSize(9);
        return $cells;
    }

    /**
     * excel Merge
     * @author atshike <atshike@foxmail.com>
     */
    private function setMergeNo($excel, $row)
    {
        $excel->mergeCells('A' . ($row - 1) . ':B' . ($row - 1))->setFontSize(9);
        static::setMergeCells($excel, 'C' . $row . ':D' . $row, 'C' . $row);
        static::setMergeCells($excel, 'E' . $row . ':F' . $row, 'E' . $row);
        static::setMergeCells($excel, 'G' . $row . ':H' . $row, 'G' . $row);
        static::setMergeCells($excel, 'I' . $row . ':J' . $row, 'I' . $row);
        static::setMergeCells($excel, 'R' . $row . ':T' . $row, 'R' . $row);
        static::setMergeCells($excel, 'U' . $row . ':Y' . $row, 'U' . $row);
    }

    /**
     * excel center
     * @author atshike <atshike@foxmail.com>
     */
    private function setMergeCells($excel, $ma, $mb, $mc = '')
    {
        $excel->mergeCells($ma)->cell($mb, function ($cells) use ($mc) {
            static::setCell($cells, $mc);
            $cells->setAlignment('center');
            $cells->setValignment('center');
        });
        return $excel;
    }

    /**
     * set Cells
     * excel
     * @author atshike <atshike@foxmail.com>
     */
    private function setCells($excel, $ma, $mc)
    {
        $excel->cell($ma, function ($cells) use ($mc) {
            $this->setCell($cells, $mc);
        });
        return $excel;
    }

    /**
     * @author atshike <atshike@foxmail.com>
     */
    private function overTimeHours($excel, $row, $rows = '')
    {
        static::setCells($excel, 'Y1', '员工签字');
        static::setMergeCells($excel, 'Y2:Y4', 'Y2', '');
        static::setMergeCells($excel, 'P1:Q1', 'P1', '姓　　名');
        static::setMergeCells($excel, 'R1:X1', 'S1', '');
        static::setMergeCells($excel, 'P2:Q2', 'P2', '员工编号');
        static::setMergeCells($excel, 'R2:X2', 'S2', '');
        static::setMergeCells($excel, 'Z2:AA2', 'Z2', '');
        static::setCells($excel, 'V3', '');
        static::setCells($excel, 'S3', '~');
        static::setCells($excel, 'X3', '');
        static::setCells($excel, 'V3', '工作时长');
        static::setCells($excel, 'Z3', '');
        static::setCells($excel, 'X3', '小时');
        static::setMergeCells($excel, 'P3:Q3', 'P3', '工作时间');
        static::setMergeCells($excel, 'P4:Q4', 'P4', '工作天数');
        static::setCells($excel, 'U4', '日');
        static::setCells($excel, 'V4', '休息时间');
        static::setCells($excel, 'Z4', '');
        static::setCells($excel, 'X4', '小时');
        static::setMergeCells($excel, 'A5:J5', 'A5', '规定工作时间');
        static::setMergeCells($excel, 'N5:N6', 'N5', '规定时间');
        static::setMergeCells($excel, 'K5:M5', 'K5', '加班');
        static::setMergeCells($excel, 'O5:O6', 'O5', '迟到早退');
        static::setMergeCells($excel, 'P5:P6', 'P5', '终差');
        static::setMergeCells($excel, 'Q5:Q6', 'Q5', '扣除');
        static::setMergeCells($excel, 'R5:T6', 'R5', '扣除理由');
        static::setMergeCells($excel, 'U5:Y6', 'U5', '备注');
        static::setMergeCells($excel, 'A6:B6', 'A6', '日期');
        static::setMergeCells($excel, 'C6:D6', 'C6', '分类');
        static::setMergeCells($excel, 'E6:F6', 'E6', '开始');
        static::setMergeCells($excel, 'G6:H6', 'G6', '结束');
        static::setMergeCells($excel, 'I6:J6', 'I6', '休息');
        static::setCells($excel, 'K6', '平日');
        static::setCells($excel, 'L6', '休息');
        static::setCells($excel, 'M6', '加班');
        static::setCells($excel, 'A' . ($row + 2), '分类：1.出勤  2.事假  3.病假');
        static::setMergeCells($excel, 'W' . ($row + 3) . ':Y' . ($row + 3), 'W' . ($row + 3), '经理签字');
        static::setMergeCells($excel, 'W' . ($row + 4) . ':W' . ($row + 6), 'W' . ($row + 4), '');
        static::setMergeCells($excel, 'X' . ($row + 4) . ':X' . ($row + 6), 'X' . ($row + 4), '');
        static::setMergeCells($excel, 'Y' . ($row + 4) . ':Y' . ($row + 6), 'Y' . ($row + 4), '');
        static::setCenter($excel, 'B1' . ':Y' . ($row + 6));
        static::setCenter($excel, 'A1');
        static::setCenter($excel, 'A5');
        static::setCenter($excel, 'A6');
        static::setCenter($excel, 'A' . ($row + 4));
        static::setCenter($excel, 'A' . ($row + 6));
        static::setCells($excel, 'E' . ($row + 5), '平日');
        static::setCells($excel, 'F' . ($row + 5), '休出');
        static::setCells($excel, 'G' . ($row + 5), '休出(日)');
        static::setCells($excel, 'H' . ($row + 5), '深夜');
    }

    /**
     * excel count 分类 excelNo()
     * @author atshike <atshike@foxmail.com>
     */
    private function tallyNum($excel, $num, $row, $text = '', $result = '')
    {
        static::setMergeCells($excel, $num . ($row + 4) . ':' . $num . ($row + 5), $num . ($row + 4), $text);
        static::setCells($excel, $num . ($row + 6), $result);
    }

    /**
     * excel count 分类
     * @author atshike <atshike@foxmail.com>
     */
    private function excelNo($excel, $row, $arr = '')
    {
        static::tallyNum($excel, 'A', $row, '年休', '');
        static::tallyNum($excel, 'B', $row, '事假', '');
        static::tallyNum($excel, 'C', $row, '病假', '');
    }

    /**
     * Set Alignment Valignment
     * excel
     * @author atshike <atshike@foxmail.com>
     */
    private function setCenter($excel, $row)
    {
        $excel->cells($row, function ($cells) {
            $cells->setAlignment('center');
            $cells->setValignment('center');
        });
    }

    /**
     * excel Set Width
     * @author atshike <atshike@foxmail.com>
     */
    private function setCellWidth($excel)
    {
        $excel->setWidth(array(
            'A' => 6,
            'B' => 6,
            'C' => 6,
            'D' => 6,
            'E' => 6,
            'F' => 6,
            'G' => 6,
            'H' => 6,
            'I' => 7,
            'J' => 6,
            'K' => 11,
            'L' => 11,
            'M' => 11,
            'N' => 11,
            'O' => 11,
            'P' => 11,
            'Q' => 11,
            'R' => 9,
            'W' => 11,
            'Y' => 11
        ));
    }
}
