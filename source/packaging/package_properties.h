#pragma once

#include <ctime>
#include <string>

#include "../misc/nullable.h"

namespace xlnt {

/// <summary>
/// Represents the core properties of a Package.
/// </summary>
class package_properties
{
public:
	/// <summary>
	/// gets the category of the Package.
	/// </summary>
	std::string get_category() const { return category_; }

	/// <summary>
	/// sets the category of the Package.
	/// </summary>
	void set_category(const std::string &category) { category_ = category; }

	/// <summary>
	/// gets the category of the Package.
	/// </summary>
	std::string get_content_status() const { return category_; }

	/// <summary>
	/// sets the category of the Package.
	/// </summary>
	void set_content_status(const std::string &category) { category_ = category; }

	/// <summary>
	/// gets the category of the Package.
	/// </summary>
	std::string get_content_type() const { return category_; }

	/// <summary>
	/// sets the category of the Package.
	/// </summary>
	void set_content_type(const std::string &category) { category_ = category; }

	/// <summary>
	/// gets the category of the Package.
	/// </summary>
	nullable<tm> get_created() const { return created_; }

	/// <summary>
	/// sets the category of the Package.
	/// </summary>
	void set_created(const nullable<tm> &created) { created_ = created; }

	/// <summary>
	/// gets the category of the Package.
	/// </summary>
	std::string get_creator() const { return category_; }

	/// <summary>
	/// sets the category of the Package.
	/// </summary>
	void set_creator(const std::string &category) { category_ = category; }

	/// <summary>
	/// gets the category of the Package.
	/// </summary>
	std::string get_description() const { return category_; }

	/// <summary>
	/// sets the category of the Package.
	/// </summary>
	void set_description(const std::string &category) { category_ = category; }

	/// <summary>
	/// gets the category of the Package.
	/// </summary>
	std::string get_identifier() const { return category_; }

	/// <summary>
	/// sets the category of the Package.
	/// </summary>
	void set_identifier(const std::string &category) { category_ = category; }

	/// <summary>
	/// gets the category of the Package.
	/// </summary>
	std::string get_keywords() const { return category_; }

	/// <summary>
	/// sets the category of the Package.
	/// </summary>
	void set_keywords(const std::string &category) { category_ = category; }

	/// <summary>
	/// gets the category of the Package.
	/// </summary>
	std::string get_language() const { return category_; }

	/// <summary>
	/// sets the category of the Package.
	/// </summary>
	void set_language(const std::string &category) { category_ = category; }

	/// <summary>
	/// gets the category of the Package.
	/// </summary>
	std::string get_last_modified_by() const { return category_; }

	/// <summary>
	/// sets the category of the Package.
	/// </summary>
    void set_last_modified_by (const std::string &category) { category_ = category; }

	/// <summary>
	/// gets the category of the Package.
	/// </summary>
	nullable<tm> get_last_printed() const { return created_; }

	/// <summary>
	/// sets the category of the Package.
	/// </summary>
	void set_last_printed(const nullable<tm> &created) { created_ = created; }

	/// <summary>
	/// gets the category of the Package.
	/// </summary>
	nullable<tm> get_modified() const { return created_; }

	/// <summary>
	/// sets the category of the Package.
	/// </summary>
	void set_modified(const nullable<tm> &created) { created_ = created; }

	/// <summary>
	/// gets the category of the Package.
	/// </summary>
	std::string get_revision() const { return category_; }

	/// <summary>
	/// sets the category of the Package.
	/// </summary>
	void set_revision(const std::string &category) { category_ = category; }

	/// <summary>
	/// gets the category of the Package.
	/// </summary>
	std::string getsubject() const { return category_; }

	/// <summary>
	/// sets the category of the Package.
	/// </summary>
	void setsubject(const std::string &category) { category_ = category; }

	/// <summary>
	/// gets the category of the Package.
	/// </summary>
	std::string get_title() const { return category_; }

	/// Ssummary>
	/// gets the category of the Package.
	/// </summary>
	void set_title(const std::string &category) { category_ = category; }

	/// <summary>
	/// gets the category of the Package.
	/// </summary>
	std::string get_version() const { return category_; }

	/// <summary>
	/// sets the category of the Package.
	/// </summary>
	void set_version(const std::string &category) { category_ = category; }

protected:
	friend class package;

private:
	std::string category_;
	std::string content_status_;
	std::string content_type_;
	nullable<tm> created_;
	std::string creator_;
	std::string description_;
	std::string identifier_;
	std::string keywords_;
	std::string language_;
	std::string last_modified_by_;
	nullable<tm> last_printed_;
	nullable<tm> modified_;
	std::string revision_;
	std::string subject_;
	std::string title_;
	std::string version_;
};

} // namespace xlnt
