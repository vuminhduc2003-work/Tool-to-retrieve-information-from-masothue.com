# Tool-to-retrieve-information-from-masothue.com

This code implements a graphical user interface (GUI) using Tkinter to process CSV files with tax codes, scrape business information from a website, and save the results into an Excel file. The key elements are:
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

CSV File Selection: The user selects a CSV file, and the application processes the tax codes by scraping data from a website (masothue.com).
Data Extraction: Using Selenium, it opens the webpage, extracts relevant business information (name, address, representative, phone number), and saves it to an Excel file.
Pause and Resume: A button allows the user to pause and resume processing.
Progress Tracking: A progress bar shows the completion percentage and estimated remaining time for processing.
Logging: The process is logged to a file (app.log), including any errors encountered during data extraction.
Multithreading: The process runs in a separate thread to avoid blocking the GUI.
Data Display: Saved data can be displayed in a new window using the show_saved_data function.
File Saving: The application saves the Excel file periodically during processing, ensuring data persistence in case of errors.
Notes and Suggestions:
Error Handling: The code includes extensive error handling, such as retrying operations or skipping tax codes when necessary.
Browser Management: Selenium’s Chrome WebDriver is managed through webdriver_manager to automatically handle ChromeDriver updates.
Path Configurations: The file save paths are hardcoded, so you might want to consider using dynamic paths or asking the user for a save location.
To implement the tool you’ve developed, the following Python libraries need to be installed and used:

1. Standard Libraries
These libraries come pre-installed with Python, so no additional installation is required:

os: For file and directory operations.
time: For time-related operations (e.g., sleep, measuring duration).
datetime: Handling date and time.
threading: To create and manage parallel processing threads.
logging: To log and track application activities.
tkinter: To create a graphical user interface (GUI).
2. External Libraries to Install
You will need to install the following libraries using pip if they are not already available in your environment:

pandas:

To work with data, especially DataFrames, making it easier to handle CSV and Excel files.
Install: pip install pandas
selenium:

To automate the web browser (in this case, Chrome) and perform web scraping tasks.
Install: pip install selenium
webdriver_manager:

Automatically manages and updates the latest version of ChromeDriver, avoiding manual installation.
Install: pip install webdriver-manager
openpyxl:

To read and write Excel files using pandas.
Install: pip install openpyxl
3. Optional Libraries
If you want to perform more complex operations with Excel files, you may need:

xlsxwriter: To create Excel files with advanced formatting features.
Install: pip install XlsxWriter

===================================================================================
= Install all: pip install pandas selenium webdriver-manager openpyxl XlsxWriter  =
===================================================================================

VIE 
Công cụ lấy thông tin từ masothue.com
Mã nguồn này triển khai một giao diện người dùng đồ họa (GUI) sử dụng Tkinter để xử lý các tệp CSV với mã số thuế, thu thập thông tin doanh nghiệp từ một trang web, và lưu kết quả vào một tệp Excel. Các thành phần chính là:
Lựa chọn tệp CSV: Người dùng chọn một tệp CSV, và ứng dụng sẽ xử lý các mã số thuế bằng cách thu thập dữ liệu từ trang web (masothue.com).
Trích xuất dữ liệu: Sử dụng Selenium, nó sẽ mở trang web, trích xuất các thông tin doanh nghiệp liên quan (tên, địa chỉ, người đại diện, số điện thoại), và lưu vào một tệp Excel.
Tạm dừng và Tiếp tục: Có một nút cho phép người dùng tạm dừng và tiếp tục quá trình xử lý.
Theo dõi tiến độ: Thanh tiến độ hiển thị phần trăm hoàn thành và thời gian còn lại ước tính cho quá trình xử lý.
Ghi log: Quá trình được ghi vào một tệp nhật ký (app.log), bao gồm cả các lỗi gặp phải khi trích xuất dữ liệu.
Đa luồng: Quá trình chạy trong một luồng riêng biệt để tránh chặn giao diện người dùng.
Hiển thị dữ liệu: Dữ liệu đã lưu có thể được hiển thị trong một cửa sổ mới bằng cách sử dụng hàm show_saved_data.
Lưu tệp: Ứng dụng lưu tệp Excel định kỳ trong quá trình xử lý, đảm bảo dữ liệu được lưu lại nếu gặp lỗi.

Ghi chú và đề xuất:

Xử lý lỗi: Mã nguồn bao gồm việc xử lý lỗi rộng rãi, chẳng hạn như thử lại các thao tác hoặc bỏ qua các mã số thuế khi cần thiết.
Quản lý trình duyệt: Trình điều khiển Chrome WebDriver của Selenium được quản lý thông qua webdriver_manager để tự động xử lý các bản cập nhật của ChromeDriver.
Cấu hình đường dẫn: Đường dẫn lưu tệp hiện đang được cố định, bạn có thể xem xét việc sử dụng đường dẫn động hoặc yêu cầu người dùng chọn vị trí lưu tệp.
Để triển khai công cụ mà bạn đã phát triển, các thư viện Python sau cần được cài đặt và sử dụng:

Thư viện tiêu chuẩn
Những thư viện này đã được cài sẵn trong Python, vì vậy không cần cài đặt bổ sung:
os: Để thao tác với tệp và thư mục.
time: Cho các thao tác liên quan đến thời gian (ví dụ: tạm dừng, đo thời lượng).
datetime: Xử lý ngày và giờ.
threading: Tạo và quản lý các luồng xử lý song song.
logging: Ghi log và theo dõi hoạt động của ứng dụng.
tkinter: Tạo giao diện người dùng đồ họa (GUI).
Thư viện bên ngoài cần cài đặt
Bạn sẽ cần cài đặt các thư viện sau bằng lệnh pip nếu chúng chưa có trong môi trường của bạn:
pandas:
Để làm việc với dữ liệu, đặc biệt là DataFrame, giúp xử lý tệp CSV và Excel dễ dàng hơn.
Cài đặt: pip install pandas

selenium:
Tự động hóa trình duyệt web (trong trường hợp này là Chrome) và thực hiện các tác vụ thu thập dữ liệu.
Cài đặt: pip install selenium

webdriver_manager:
Quản lý và cập nhật tự động phiên bản mới nhất của ChromeDriver, tránh việc cài đặt thủ công.
Cài đặt: pip install webdriver-manager

openpyxl:
Để đọc và ghi tệp Excel bằng pandas.
Cài đặt: pip install openpyxl

Thư viện tùy chọn
Nếu bạn muốn thực hiện các thao tác phức tạp hơn với tệp Excel, bạn có thể cần:
xlsxwriter:
Tạo các tệp Excel với các tính năng định dạng nâng cao.
Cài đặt: pip install XlsxWriter

