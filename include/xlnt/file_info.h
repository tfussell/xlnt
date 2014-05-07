#pragma once

#include <ctime>
#include <string>
#include <tuple>

namespace xlnt {

/// <summary>
/// Provides attributes for files and directories.
/// </summary>
enum class file_attributes
{
	/// <summary>
	/// The file is a candidate for backup or removal.
	/// </summary>
	Archive,
	/// <summary>
	/// The file is compressed.
	/// </summary>
	Compressed,
	/// <summary>
	/// Reserved for future use.
	/// </summary>
	Device,
	/// <summary>
	/// The file is a directory.
	/// </summary>
	Directory,
	/// <summary>
	/// The file or directory is encrypted. For a file, this means that all data in the file is encrypted. For a directory, this means that encryption is the default for newly created files and directories.
	/// </summary>
	Encrypted,
	/// <summary>
	/// The file is hidden, and thus is not included in an ordinary directory listing.
	/// </summary>
	Hidden,
	/// <summary>
	/// The file or directory includes data integrity support. When this value is applied to a file, all data streams in the file have integrity support. When this value is applied to a directory, all new files and subdirectories within that directory, by default, include integrity support.
	/// </summary>
	IntegrityStream,
	/// <summary>
	/// The file is a standard file that has no special attributes. This attribute is valid only if it is used alone.
	/// </summary>
	Normal,
	/// <summary>
	/// The file or directory is excluded from the data integrity scan. When this value is applied to a directory, by default, all new files and subdirectories within that directory are excluded from data integrity.
	/// </summary>
	NoScrubData,
	/// <summary>
	/// The file will not be indexed by the operating system's content indexing service.
	/// </summary>
	NotContentIndexed,
	/// <summary>
	/// The file is offline. The data of the file is not immediately available.
	/// </summary>
	Offline,
	/// <summary>
	/// The file is read-only.
	/// </summary>
	ReadOnly,
	/// <summary>
	/// The file contains a reparse point, which is a block of user-defined data associated with a file or a directory.
	/// </summary>
	ReparsePoint,
	/// <summary>
	/// The file is a sparse file. Sparse files are typically large files whose data consists of mostly zeros.
	/// </summary>
	SparseFile,
	/// <summary>
	/// The file is a system file. That is, the file is part of the operating system or is used exclusively by the operating system.
	/// </summary>
	System,
	/// <summary>
	/// The file is temporary. A temporary file contains data that is needed while an application is executing but is not needed after the application is finished. File systems try to keep all the data in memory for quicker access rather than flushing the data back to mass storage. A temporary file should be deleted by the application as soon as it is no longer needed.
	/// </summary>
	Temporary
};

class file_info
{
public:
	file_info(const std::string &fileName);

	file_attributes get_attributes() const { return attributes_; }
	void set_attributes(file_attributes attributes) { attributes_ = attributes; }

	tm get_creation_time() const { return creationTime_; }
	void set_creation_time(const tm &creationTime) { creationTime_ = creationTime; }

	tm get_creation_time_utc() const { return time_to_utc(creationTime_); }
	void set_creation_time_utc(const tm &creationTime) { creationTime_ = time_from_utc(creationTime); }

	std::string get_directory_name() const { return directoryName_; }

	bool exists() const;

	std::string get_extension() const { return split_extension(fullName_).second; }

	std::string get_full_name() const { return fullName_; }

	bool IsReadOnly() const { return readOnly_; }
	void set_ReadOnly(bool readOnly) { readOnly_ = readOnly; }

	tm get_LastAccessTime() const { return creationTime_; }
	void set_LastAccessTime(const tm &creationTime) { creationTime_ = creationTime; }

	tm get_LastAccessUtc() const { return time_to_utc(creationTime_); }
	void set_LastAccessUtc(const tm &creationTime) { creationTime_ = time_from_utc(creationTime); }

	tm get_LastWriteTime() const { return creationTime_; }
	void set_LastWriteTime(const tm &creationTime) { creationTime_ = creationTime; }

	tm get_LastWriteTimeUtc() const { return time_to_utc(creationTime_); }
	void set_LastWriteTimeUtc(const tm &creationTime) { creationTime_ = time_from_utc(creationTime); }

	void delete_();
	void encrypt();
	void move_to();
	void open();
	void open_read();
	void open_text();
	void open_write();
	void refresh();
	void replace(const std::string &, const std::string &);
	void replace(const std::string &, const std::string &, bool);
	void set_access_control();
	std::string to_string();

private:
	static tm time_to_utc(const tm &time);
	static tm time_from_utc(const tm &time);
	static std::pair<std::string, std::string> split_extension(const std::string &name);

	file_attributes attributes_;
	tm creationTime_;
	std::string directoryName_;
	bool exists_;
	std::string fullName_;
	bool readOnly_;
	tm lastAccessTime_;
	tm lastWriteTime_;
	int length_;

	std::string path_;
};

} // namespace xlnt
