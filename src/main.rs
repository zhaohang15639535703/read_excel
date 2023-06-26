#![allow(unused)]
#![allow(non_snake_case)]

use calamine::{open_workbook_auto, DataType, Error, RangeDeserializerBuilder, Reader};

#[derive(Default, Debug)]
struct ProcessArgument {
    arrivalTime: u32,
    burstPeriod: u32,
    priority: u32,
}
fn main() -> Result<(), Error> {
    let path = format!("{}/tests/test.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut workbook = open_workbook_auto(path)?;
    let range = workbook
        .worksheet_range("Sheet1")
        .ok_or(Error::Msg("Cannot find 'Sheet1'"))??;
    // 获取按行读取的迭代器
    let mut iterator = range.rows();
    // 先读取第一行,记录表头所在每一列的index
    let first_row = iterator.next().ok_or(Error::Msg("no first row"))?;
    let (mut arrival_time_row, mut burst_period_row, mut priority_row) =
        (0_usize, 0_usize, 0_usize);
    let mut count = 0;
    for (index, col) in first_row.iter().enumerate() {
        if let Some(val) = col.get_string() {
            if val.eq_ignore_ascii_case("arrivalTime") {
                arrival_time_row = index;
                count += 1;
            } else if val.eq_ignore_ascii_case("burstPeriod") {
                burst_period_row = index;
                count += 1;
            } else if val.eq_ignore_ascii_case("priority") {
                priority_row = index;
                count += 1;
            }
        } else {
            //忽略其他列
            continue;
        }
    }
    // 缺少参数
    if (count != 3) {
        return Err(Error::Msg("misssing arguments"));
    }
    println!("arrival_time_row: {}", arrival_time_row);
    println!("burst_period_row: {}", burst_period_row);
    println!("priority_row: {}", priority_row);

    //从第二行开始读取,遇到不合法的数字记录下来,然后迭代下一个
    let mut process_arguments_array: Vec<ProcessArgument> = vec![];
    let mut skip = 0;
    for row in iterator {
        let mut process_argument = ProcessArgument::default();

        process_argument.arrivalTime = match row.get(arrival_time_row) {
            Some(val) => {
                if let Ok(num) = val.to_string().parse::<u32>() {
                    num
                } else {
                    //该行该列此数据转换失败
                    skip += 1;
                    continue;
                }
            }
            None => {
                // 获取该行该列数据失败
                skip += 1;
                continue;
            }
        };
        process_argument.burstPeriod = match row.get(burst_period_row) {
            Some(val) => {
                if let Ok(num) = val.to_string().parse::<u32>() {
                    num
                } else {
                    //该行该列此数据转换失败
                    skip += 1;
                    continue;
                }
            }
            None => {
                skip += 1;
                continue;
            }
        };
        process_argument.priority = match row.get(priority_row) {
            Some(val) => {
                if let Ok(num) = val.to_string().parse::<u32>() {
                    num
                } else {
                    //该行该列此数据转换失败
                    skip += 1;
                    continue;
                }
            }
            None => {
                skip += 1;
                continue;
            }
        };
        process_arguments_array.push(process_argument);
    }
    println!("{:#?}", process_arguments_array);
    println!("skip: {}", skip);
    Ok(())
}
