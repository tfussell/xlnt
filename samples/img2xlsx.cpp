// Copyright (c) 2017-2018 Thomas Fussell
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, WRISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#include <array>
#include <iostream>
#include <unordered_set>
#include <vector>

#include <xlnt/xlnt.hpp>

#define STB_IMAGE_IMPLEMENTATION
#include "stb_image.h"

namespace {

// This sample demonstrates the use of some complex formatting methods to create
// a workbook in which each cell has a fill based on the pixels of an image
// thereby appearing as a mosaic of the given image. Two methods for achieving
// this effect are demonstrated: cell-level format and conditional formatting.

// Clean up the code with some alias declarations
using byte = std::uint8_t;
using pixel = std::array<byte, 3>;
using pixmap = std::vector<std::vector<pixel>>;

// Specify whether the CF or cell format method should be used
static const bool use_conditional_formatting = false;

// Returns a 2D vector of pixels from a given filename. Accepts all file types
// supported by stb_image (JPEG, PNG, TGA, BMP, PSD, GIF, HDR, PIC, PNM).
pixmap load_image(const std::string &filename)
{
	int width = 0;
	int height = 0;
	int bpp = 0;

	// Must be freed after use with stbi_image_free
	auto image_data = stbi_load(filename.c_str(), &width, &height, &bpp, 0);

	if (image_data == nullptr)
	{
			throw xlnt::exception("bad image or file not found: " + filename);
	}

	pixmap result;

	for (auto row = 0; row < height; ++row)
	{
		std::vector<pixel> row_pixels;

		for (auto column = 0; column < width; ++column)
		{
			auto r = image_data[row * width * bpp + column * bpp];
			auto g = image_data[row * width * bpp + column * bpp + 1];
			auto b = image_data[row * width * bpp + column * bpp + 2];
			auto current_pixel = pixel({ r, g, b });
			row_pixels.push_back(current_pixel);
		}

		result.push_back(row_pixels);
	}

	stbi_image_free(image_data);

	return result;
}

// Builds and returns a workbook in which each cell has a value
// equal to the color of the pixel in the given image. A conditional format
// is created for every color to set the background fill of each cell to the
// color of its value. This is very slow for large images but is intended
// to illustrate a realistic use of conditional formatting.
xlnt::workbook build_workbook_cf(const pixmap &image)
{
	// Create a default workbook with a single worksheet
	xlnt::workbook wb;
	// Get the active sheet (the only sheet in this case)
	auto ws = wb.active_sheet();
	// The reference to the cell which is being operated upon
	auto current_cell = xlnt::cell_reference("A1");
	// The range of cells which will be modified. This is required for conditional formats
	auto range = ws.range(xlnt::range_reference(1, 1,
        static_cast<xlnt::column_t::index_t>(image[0].size()),
        static_cast<xlnt::row_t>(image.size())));

	// Track the previously created conditonal formats so they are only created once
	std::unordered_set<std::string> defined_colors;

	// Iterate over each row in the source image
	for (const auto &image_row : image)
	{
		// Iterate over each pixel in the image row
		for (const auto &image_pixel : image_row)
		{
			// Build an xlnt compatible RGB color from the pixel byte array
			const auto color = xlnt::rgb_color(image_pixel[0], image_pixel[1], image_pixel[2]);

			// Only create the conditional format if it doesn't yet exist
			if (defined_colors.count(color.hex_string()) == 0)
			{
				// The condition under which the conditional format applies to a cell
				// In this case, the condition is satisfied when the text of the cell
				// contains the hex string representing the pixel color.
				const auto condition = xlnt::condition::text_contains(color.hex_string());
				// Create a new conditional format with the above condition on the image pixel range
				auto format = range.conditional_format(condition);
				// Define the fill for the conditional format
				format.fill(xlnt::pattern_fill().background(color));

				// Record the created of this CF
				defined_colors.insert(color.hex_string());
			}

			// Dereference the cell at the current position and set its value to the pixel color
			ws.cell(current_cell).value(color.hex_string());
			// Increment the column
			current_cell.column_index(current_cell.column_index() + 1);
		}

		// Reached the end of the row--move to the first column of the next row
		current_cell.row(current_cell.row() + 1);
		current_cell.column_index("A");

		// Show some progress, it can take a while...
		std::cout << current_cell.row() << " " << defined_colors.size() << std::endl;
	}

	// Return the resulting workbook
	return wb;
}

// Builds and returns a workbook in which each cell has a value
// equal to the color of the pixel in the given image. To accomplish this,
// a named style is created for every color in the image and is applied
// to each cell corresponding to pixels of that color.
xlnt::workbook build_workbook_normal(const pixmap &image)
{
	// Create a default workbook with a single worksheet
	xlnt::workbook wb;
	// Get the active sheet (the only sheet in this case)
	auto ws = wb.active_sheet();
	// The reference to the cell which is being operated upon
	auto current_cell = xlnt::cell_reference("A1");

	// Iterate over each row in the source image
	for (const auto &image_row : image)
	{
		// Iterate over each pixel in the image row
		for (const auto &image_pixel : image_row)
		{
			// Build an xlnt compatible RGB color from the pixel byte array
			const auto color = xlnt::rgb_color(image_pixel[0], image_pixel[1], image_pixel[2]);

			// Only create the style if it doesn't yet exist in the workbook
			if (!wb.has_style(color.hex_string()))
			{
				// A style is constructed on a workbook by providing a unique name
				auto color_style = wb.create_style(color.hex_string());
				// Set the fill to a solid fill of the pixel color
				color_style.fill(xlnt::fill::solid(color));
			}

			// Dereference the cell at the current position and use the previously created color style
			ws.cell(current_cell).style(color.hex_string());
			// Increment the column
			current_cell.column_index(current_cell.column_index() + 1);
		}

		// Reached the end of the row--move to the first column of the next row
		current_cell.row(current_cell.row() + 1);
		current_cell.column_index("A");

		// Show some progress, it can take a while...
		std::cout << current_cell.row() << std::endl;
	}

	// Return the resulting workbook
	return wb;
}

// Builds and returns a workbook in which each cell has a value
// equal to the color of the pixel in the given image using the
// conditional formatting or cell-level formatting depending on
// the value of use_conditional_formatting defined at the top of
// this file.
xlnt::workbook build_workbook(const pixmap &image)
{
	if (use_conditional_formatting)
	{
		return build_workbook_cf(image);
	}
	else
	{
		return build_workbook_normal(image);
	}
}

} // namespace

// Entry point
int main(int argc, char *argv[])
{
	// Ensure that there is a correct number of arguments
	if (argc < 3)
	{
		std::cout << "usage: img2xlsx <input> <output>" << std::endl;
		return 0;
	}

	// The first argument is the name of the input image
	const auto input = std::string(argv[1]);
	// Load the input image. An exception will be thrown if it doesn't exist.
	const auto image = load_image(input);
	// Build the workbook from the image
	auto workbook = build_workbook(image);
	// The second argument is the name of the file to save the workbook as.
	const auto output = std::string(argv[2]);
	// Save the workbook
	workbook.save(output);

    return 0;
}
