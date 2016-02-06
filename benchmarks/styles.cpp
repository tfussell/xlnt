#include <iterator>
#include <random>
#include <xlnt/xlnt.hpp>

template<typename Iter>
Iter random_choice(Iter start, Iter end) {
    static std::random_device rd;
    static std::mt19937 gen(rd());

    std::uniform_int_distribution<> dis(0, std::distance(start, end) - 1);
    std::advance(start, dis(gen));

    return start;
}

std::vector<xlnt::style> generate_all_styles()
{
    std::vector<xlnt::style> styles;

    std::vector<xlnt::vertical_alignment> vertical_alignments = {xlnt::vertical_alignment::center, xlnt::vertical_alignment::justify, xlnt::vertical_alignment::top, xlnt::vertical_alignment::bottom};
    std::vector<xlnt::horizontal_alignment> horizontal_alignments = {xlnt::horizontal_alignment::center, xlnt::horizontal_alignment::center_continuous, xlnt::horizontal_alignment::general, xlnt::horizontal_alignment::justify, xlnt::horizontal_alignment::left, xlnt::horizontal_alignment::right};
    std::vector<std::string> font_names = {"Calibri", "Tahoma", "Arial", "Times New Roman"};
    std::vector<int> font_sizes = {11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35};
    std::vector<bool> bold_options = {true, false};
    std::vector<xlnt::font::underline_style> underline_options = {xlnt::font::underline_style::single, xlnt::font::underline_style::none};
    std::vector<bool> italic_options = {true, false};
    
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
                                xlnt::style s;

                                xlnt::font f;
                                f.set_name(name);
                                f.set_size(size);
                                f.set_italic(italic);
                                f.set_underline(underline);
                                f.set_bold(bold);
                                s.set_font(f);

                                xlnt::alignment a;
                                a.set_vertical(vertical_alignment);
                                a.set_horizontal(horizontal_alignment);
                                s.set_alignment(a);

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

xlnt::workbook optimized_workbook(const std::vector<xlnt::style> &styles, int n)
{
    xlnt::workbook wb;
    wb.set_optimized_write(true);
    auto worksheet = wb.create_sheet();

    for(int i = 1; i < n; i++)
    {
        auto style = *random_choice(styles.begin(), styles.end());
        worksheet.append({{0, style}});
    }

    return wb;
}

xlnt::workbook non_optimized_workbook(const std::vector<xlnt::style> &styles, int n)
{
    xlnt::workbook wb;

    for(int idx = 1; idx < n; idx++)
    {
        auto worksheet = *random_choice(wb.begin(), wb.end());
        auto cell = worksheet.get_cell({1, (xlnt::row_t)idx + 1});
        cell.set_value(0);
        cell.set_style(*random_choice(styles.begin(), styles.end()));
    }

    return wb;
}

void to_profile(xlnt::workbook &wb, const std::string &f, int n)
{
    auto t = 0;//-time.time();
    wb.save(f);
    std::cout << "took " << t << "s for " << n << " styles";
}

int main()
{
    auto styles = generate_all_styles();
    int n = 10000;

    for(auto func : {&optimized_workbook, &non_optimized_workbook})
    {
	std::cout << (func == &optimized_workbook ? "optimized_workbook" : "non_optimized_workbook") << std::endl;
	auto wb = func(styles, n);
	std::string f = "/tmp/xlnt.xlsx";
	to_profile(wb, f, n);
    }
}
