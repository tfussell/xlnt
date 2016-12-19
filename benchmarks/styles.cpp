#include <chrono>
#include <iostream>
#include <iterator>
#include <random>
#include <xlnt/xlnt.hpp>

std::size_t current_time()
{
    return std::chrono::duration<double, std::milli>(std::chrono::system_clock::now().time_since_epoch()).count();
}

std::size_t random_index(std::size_t max)
{
    static std::random_device rd;
    static std::mt19937 gen(rd());

    std::uniform_int_distribution<> dis(0, max - 1);

    return dis(gen);
}

std::vector<xlnt::style> generate_all_styles(xlnt::workbook &wb)
{
    std::vector<xlnt::style> styles;

    std::vector<xlnt::vertical_alignment> vertical_alignments = {xlnt::vertical_alignment::center, xlnt::vertical_alignment::justify, xlnt::vertical_alignment::top, xlnt::vertical_alignment::bottom};
    std::vector<xlnt::horizontal_alignment> horizontal_alignments = {xlnt::horizontal_alignment::center, xlnt::horizontal_alignment::center_continuous, xlnt::horizontal_alignment::general, xlnt::horizontal_alignment::justify, xlnt::horizontal_alignment::left, xlnt::horizontal_alignment::right};
    std::vector<std::string> font_names = {"Calibri", "Tahoma", "Arial", "Times New Roman"};
    std::vector<int> font_sizes = {11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35};
    std::vector<bool> bold_options = {true, false};
    std::vector<xlnt::font::underline_style> underline_options = {xlnt::font::underline_style::single, xlnt::font::underline_style::none};
    std::vector<bool> italic_options = {true, false};
    std::size_t index = 0;
    
    for(auto vertical_alignment : vertical_alignments)
    {
        for(auto horizontal_alignment : horizontal_alignments)
        {
            for(auto name : font_names)
            {
                for(auto size : font_sizes)
                {
                    for(auto bold : bold_options)
                    {
                        for(auto underline : underline_options)
                        {
                            for(auto italic : italic_options)
                            {
                                auto s = wb.create_style(std::to_string(index++));

                                xlnt::font f;
                                f.name(name);
                                f.size(size);
                                f.italic(italic);
                                f.underline(underline);
                                f.bold(bold);
                                s.font(f);

                                xlnt::alignment a;
                                a.vertical(vertical_alignment);
                                a.horizontal(horizontal_alignment);
                                s.alignment(a);

                                styles.push_back(s);
                            }
                        }
                    }
                }
            }
        }
    }

    return styles;
}

xlnt::workbook non_optimized_workbook(int n)
{
    xlnt::workbook wb;
    auto styles = generate_all_styles(wb);

    for(int idx = 1; idx < n; idx++)
    {
        auto worksheet = wb[random_index(wb.sheet_count())];
        auto cell = worksheet.cell(1, (xlnt::row_t)idx);
        cell.value(0);
        cell.style(styles.at(random_index(styles.size())));
    }

    return wb;
}

void to_profile(xlnt::workbook &wb, const std::string &f, int n)
{
    auto start = current_time();
    wb.save(f);
    auto elapsed = current_time() - start;
    std::cout << "took " << elapsed / 1000.0 << "s for " << n << " styles" << std::endl;
}

int main()
{
    int n = 10000;
    auto wb = non_optimized_workbook(n);
    std::string f = "temp.xlsx";
    to_profile(wb, f, n);

    return 0;
}
