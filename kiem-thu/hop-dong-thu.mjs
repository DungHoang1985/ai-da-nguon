// Hợp đồng thử cho Đợt 1A: khoảng 50 trang, đánh số 3 cấp do Word tự đánh ("Điều 1." / "1.1." / "a)") gắn kiểu Heading 1–3,
// thêm một phụ lục đánh số GÕ TAY, 30 sửa đổi được theo dõi và 5 nhận xét của tác giả "Đối tác".
// Nội dung bịa để thử (tên pháp nhân, số liệu đều không có thật).
//   node kiem-thu/hop-dong-thu.mjs <tệp ra> [số câu đệm mỗi khoản, mặc định 10]
import path from "node:path";
import { fileURLToPath } from "node:url";
import { para, run, del, ins, commented, buildDocx } from "./docx-thu.mjs";

// Mỗi điều: [tiêu đề, [khoản…]]; khoản là chuỗi, hoặc [chuỗi, [điểm…]]
const DIEU = [
  ["Định nghĩa và giải thích", [
    "“Khu Nhà Xưởng” là toàn bộ nhà xưởng, văn phòng và công trình phụ trợ tại lô CN-07, có tổng diện tích sàn 12.600 m² như mô tả tại Phụ lục 01.",
    "“Ngày Bàn Giao” là ngày hai bên ký biên bản bàn giao Khu Nhà Xưởng theo Điều 8.",
    "“Tiền Thuê” là khoản tiền Bên Thuê trả cho Bên Cho Thuê để sử dụng Khu Nhà Xưởng, tính theo Điều 4.",
    "“Ngày Làm Việc” là các ngày từ thứ Hai đến thứ Sáu, trừ ngày nghỉ lễ theo quy định.",
  ]],
  ["Đối tượng của hợp đồng", [
    "Bên Cho Thuê đồng ý cho thuê và Bên Thuê đồng ý thuê Khu Nhà Xưởng để sản xuất linh kiện cơ khí chính xác và làm văn phòng điều hành.",
    "Diện tích thuê được xác định theo số đo thực tế tại Ngày Bàn Giao; chênh lệch dưới 1% so với diện tích nêu tại Phụ lục 01 không làm thay đổi Tiền Thuê.",
    "Bên Thuê không được dùng Khu Nhà Xưởng vào mục đích khác nếu chưa có chấp thuận bằng văn bản của Bên Cho Thuê.",
  ]],
  ["Thời hạn thuê", [
    "Thời hạn thuê là 10 năm kể từ Ngày Bàn Giao.",
    "Bên Thuê được ưu tiên gia hạn thêm 05 năm nếu gửi đề nghị bằng văn bản chậm nhất 09 tháng trước ngày hết hạn.",
    "Hết thời hạn thuê mà hai bên không gia hạn thì hợp đồng chấm dứt và Bên Thuê hoàn trả mặt bằng theo Điều 24.",
  ]],
  ["Tiền thuê", [
    "Tiền Thuê được tính theo diện tích sàn thực tế và đơn giá quy định tại Điều này.",
    "Đơn giá thuê là 450.000 đồng cho mỗi mét vuông mỗi tháng, chưa gồm thuế giá trị gia tăng, áp dụng cố định trong ba năm đầu.",
    "Từ năm thứ tư, đơn giá thuê tăng 5% sau mỗi hai năm so với đơn giá của kỳ liền trước.",
    "Tiền Thuê không bao gồm phí quản lý, tiền điện, nước, xử lý nước thải và các dịch vụ khác do Bên Thuê sử dụng.",
  ]],
  ["Đặt cọc", [
    "Bên Thuê đặt cọc cho Bên Cho Thuê một khoản tiền bằng 03 tháng Tiền Thuê trong vòng 10 Ngày Làm Việc kể từ ngày ký hợp đồng.",
    "Tiền đặt cọc không tính lãi và được hoàn trả trong vòng 30 ngày kể từ ngày Bên Thuê hoàn thành mọi nghĩa vụ khi chấm dứt hợp đồng.",
    ["Bên Cho Thuê được khấu trừ tiền đặt cọc trong các trường hợp sau:", [
      "Bên Thuê chậm trả Tiền Thuê quá thời hạn quy định;",
      "Bên Thuê gây hư hỏng cho Khu Nhà Xưởng mà không khắc phục;",
      "Bên Thuê đơn phương chấm dứt hợp đồng trái thỏa thuận.",
    ]],
  ]],
  ["Thanh toán", [
    "Tiền Thuê được thanh toán theo quý, trong vòng 15 ngày đầu của mỗi quý, bằng chuyển khoản vào tài khoản của Bên Cho Thuê.",
    "Bên Cho Thuê xuất hóa đơn giá trị gia tăng cho từng kỳ thanh toán trước ngày đến hạn ít nhất 07 ngày.",
    "Nếu chậm thanh toán, Bên Thuê phải trả lãi chậm trả 0,05% mỗi ngày trên số tiền chậm trả, tính từ ngày đến hạn tới ngày trả đủ.",
    "Chậm thanh toán quá 45 ngày được coi là vi phạm nghiêm trọng và Bên Cho Thuê có quyền chấm dứt hợp đồng theo Điều 22.",
  ]],
  ["Phí quản lý và dịch vụ", [
    "Phí quản lý khu công nghiệp là 12.000 đồng cho mỗi mét vuông mỗi tháng và được điều chỉnh theo thông báo của đơn vị quản lý hạ tầng.",
    "Tiền điện, nước và xử lý nước thải được tính theo chỉ số công tơ và đơn giá của nhà cung cấp tại thời điểm sử dụng.",
    "Bên Thuê tự ký hợp đồng viễn thông, vệ sinh công nghiệp và thu gom chất thải với nhà cung cấp do mình lựa chọn.",
  ]],
  ["Bàn giao mặt bằng", [
    "Bên Cho Thuê bàn giao Khu Nhà Xưởng cho Bên Thuê chậm nhất vào ngày 01 tháng 03 năm 2027.",
    "Tại thời điểm bàn giao, hai bên lập biên bản ghi nhận hiện trạng, chỉ số công tơ và danh mục thiết bị kèm theo.",
    "Nếu Bên Cho Thuê chậm bàn giao quá 30 ngày, Bên Thuê được miễn Tiền Thuê tương ứng với gấp đôi số ngày chậm.",
  ]],
  ["Sử dụng mặt bằng", [
    "Bên Thuê sử dụng Khu Nhà Xưởng đúng công năng thiết kế và đúng tải trọng sàn cho phép là 2 tấn trên mỗi mét vuông.",
    "Bên Thuê tuân thủ nội quy khu công nghiệp do đơn vị quản lý hạ tầng ban hành.",
    "Bên Thuê phải thông báo cho Bên Cho Thuê trước khi đưa vào Khu Nhà Xưởng hóa chất hoặc vật liệu dễ cháy nổ.",
  ]],
  ["Cải tạo và lắp đặt", [
    "Bên Thuê được lắp đặt máy móc, vách ngăn và hệ thống phụ trợ phục vụ sản xuất bằng chi phí của mình.",
    "Mọi cải tạo làm thay đổi kết cấu chịu lực, mái hoặc hệ thống phòng cháy chữa cháy phải được Bên Cho Thuê chấp thuận bằng văn bản trước khi thi công.",
    "Bên Cho Thuê trả lời đề nghị cải tạo trong vòng 15 Ngày Làm Việc; quá thời hạn này mà không trả lời thì coi như đồng ý.",
  ]],
  ["Bảo trì và sửa chữa", [
    "Bên Cho Thuê chịu trách nhiệm sửa chữa kết cấu chính, mái, tường bao và hạ tầng kỹ thuật bên ngoài nhà xưởng.",
    "Bên Thuê chịu trách nhiệm bảo trì thường xuyên phần bên trong, thiết bị do mình lắp đặt và hệ thống chiếu sáng.",
    "Hư hỏng thuộc trách nhiệm của Bên Cho Thuê phải được khắc phục trong vòng 07 ngày kể từ ngày nhận thông báo; trường hợp khẩn cấp là 24 giờ.",
  ]],
  ["Bảo hiểm", [
    "Bên Cho Thuê mua bảo hiểm mọi rủi ro tài sản cho phần kết cấu của Khu Nhà Xưởng.",
    "Bên Thuê mua bảo hiểm cho máy móc, hàng hóa và tài sản của mình đặt tại Khu Nhà Xưởng.",
    "Mỗi bên mua bảo hiểm trách nhiệm đối với bên thứ ba cho hoạt động của mình.",
    "Bên Thuê mua bảo hiểm cháy nổ bắt buộc với mức trách nhiệm không thấp hơn 80 tỷ đồng và gửi bản sao giấy chứng nhận cho Bên Cho Thuê trước Ngày Bàn Giao.",
  ]],
  ["Phòng cháy chữa cháy và an toàn", [
    "Bên Cho Thuê bàn giao Khu Nhà Xưởng đã được nghiệm thu về phòng cháy chữa cháy theo quy định.",
    "Bên Thuê duy trì đội phòng cháy chữa cháy cơ sở và tổ chức diễn tập ít nhất một lần mỗi năm.",
    "Bên Thuê chịu trách nhiệm về an toàn lao động đối với người lao động và nhà thầu của mình.",
  ]],
  ["Bảo vệ môi trường", [
    "Bên Thuê thực hiện đầy đủ thủ tục môi trường cho hoạt động sản xuất của mình trước khi vận hành.",
    "Nước thải phải được xử lý sơ bộ đạt tiêu chuẩn đấu nối của khu công nghiệp trước khi xả vào hệ thống chung.",
    "Bên Thuê chịu toàn bộ chi phí khắc phục nếu hoạt động của mình gây ô nhiễm.",
  ]],
  ["Cho thuê lại và chuyển nhượng", [
    "Bên Thuê không được cho thuê lại toàn bộ hoặc một phần Khu Nhà Xưởng, trừ khi cho công ty con mà mình nắm trên 50% vốn điều lệ thuê lại và đã thông báo trước 30 ngày.",
    "Bên Thuê không được chuyển nhượng quyền và nghĩa vụ theo hợp đồng nếu chưa có chấp thuận bằng văn bản của Bên Cho Thuê.",
    "Bên Cho Thuê được chuyển nhượng Khu Nhà Xưởng cho bên thứ ba với điều kiện bên nhận chuyển nhượng kế thừa toàn bộ hợp đồng này.",
  ]],
  ["Quyền của Bên Cho Thuê", [
    "Nhận đủ Tiền Thuê và các khoản phải thu khác đúng hạn.",
    "Kiểm tra việc sử dụng Khu Nhà Xưởng sau khi báo trước ít nhất 02 Ngày Làm Việc, trừ trường hợp khẩn cấp.",
    "Yêu cầu Bên Thuê bồi thường thiệt hại do lỗi của Bên Thuê gây ra.",
  ]],
  ["Nghĩa vụ của Bên Cho Thuê", [
    "Bàn giao Khu Nhà Xưởng đúng thời hạn, đúng hiện trạng đã thỏa thuận.",
    "Bảo đảm quyền sử dụng ổn định của Bên Thuê trong suốt thời hạn thuê.",
    "Hỗ trợ Bên Thuê trong các thủ tục xin giấy phép liên quan đến việc sử dụng Khu Nhà Xưởng.",
  ]],
  ["Quyền của Bên Thuê", [
    "Sử dụng Khu Nhà Xưởng và hạ tầng dùng chung của khu công nghiệp trong suốt thời hạn thuê.",
    ["Yêu cầu Bên Cho Thuê giảm Tiền Thuê trong các trường hợp sau:", [
      "Khu Nhà Xưởng bị hư hỏng không do lỗi của Bên Thuê làm giảm diện tích sử dụng;",
      "nguồn điện hoặc nước bị gián đoạn quá 72 giờ liên tục do lỗi của Bên Cho Thuê;",
      "đường vào khu công nghiệp bị phong tỏa quá 05 ngày.",
    ]],
    "Gắn biển hiệu tại vị trí được đơn vị quản lý hạ tầng chấp thuận.",
  ]],
  ["Nghĩa vụ của Bên Thuê", [
    "Trả Tiền Thuê và các khoản phí đầy đủ, đúng hạn.",
    "Giữ gìn Khu Nhà Xưởng và sửa chữa hư hỏng do mình gây ra.",
    "Không làm ảnh hưởng đến hoạt động của các doanh nghiệp lân cận.",
  ]],
  ["Cam đoan và bảo đảm", [
    "Mỗi bên cam đoan được thành lập hợp pháp và có đủ thẩm quyền ký kết, thực hiện hợp đồng này.",
    "Bên Cho Thuê cam đoan Khu Nhà Xưởng không có tranh chấp, không bị kê biên và không thế chấp cho bên thứ ba ngoài ngân hàng đã nêu tại Phụ lục 01.",
    "Bên Thuê cam đoan nguồn tiền thanh toán là hợp pháp.",
  ]],
  ["Thông báo vi phạm và khắc phục", [
    "Bên bị vi phạm gửi thông báo bằng văn bản nêu rõ hành vi vi phạm và yêu cầu khắc phục ngay lập tức.",
    "Bên vi phạm có 30 ngày kể từ ngày nhận thông báo để khắc phục, trừ trường hợp hợp đồng quy định thời hạn khác.",
    "Trong thời gian khắc phục, các nghĩa vụ khác của hai bên vẫn tiếp tục được thực hiện.",
  ]],
  ["Chấm dứt hợp đồng", [
    "Hợp đồng chấm dứt khi hết thời hạn thuê mà không gia hạn, hoặc theo thỏa thuận bằng văn bản của hai bên.",
    "Bên Cho Thuê được đơn phương chấm dứt hợp đồng nếu Bên Thuê chậm thanh toán quá thời hạn nêu tại Điều 6 hoặc sử dụng sai mục đích mà không khắc phục.",
    "Bên Thuê được đơn phương chấm dứt hợp đồng nếu Bên Cho Thuê chậm bàn giao quá 90 ngày hoặc Khu Nhà Xưởng không thể sử dụng quá 60 ngày liên tục.",
    "Bên Thuê được chấm dứt hợp đồng trước hạn không cần lý do sau năm thứ năm, với điều kiện báo trước 12 tháng và trả khoản phí chấm dứt bằng 06 tháng Tiền Thuê.",
  ]],
  ["Phạt vi phạm và bồi thường", [
    "Bên vi phạm nghĩa vụ hợp đồng chịu phạt 8% giá trị phần nghĩa vụ bị vi phạm.",
    "Ngoài tiền phạt, bên vi phạm bồi thường toàn bộ thiệt hại thực tế, trực tiếp cho bên kia.",
    "Tổng mức bồi thường của mỗi bên trong một năm không vượt quá 12 tháng Tiền Thuê, trừ trường hợp cố ý vi phạm.",
  ]],
  ["Hoàn trả mặt bằng", [
    "Khi chấm dứt hợp đồng, Bên Thuê hoàn trả Khu Nhà Xưởng trong vòng 30 ngày theo hiện trạng lúc nhận bàn giao, trừ hao mòn tự nhiên.",
    "Bên Thuê tháo dỡ máy móc, thiết bị do mình lắp đặt và khắc phục các vị trí đã cải tạo, trừ khi Bên Cho Thuê đồng ý giữ lại.",
    "Chậm hoàn trả mặt bằng, Bên Thuê trả tiền sử dụng bằng 150% Tiền Thuê tính theo ngày cho thời gian chậm.",
  ]],
  ["Bảo mật", [
    "Mỗi bên giữ bí mật nội dung hợp đồng và thông tin kinh doanh của bên kia trong thời hạn thuê và 03 năm sau đó.",
    "Nghĩa vụ bảo mật không áp dụng với thông tin phải cung cấp theo yêu cầu của cơ quan nhà nước có thẩm quyền.",
  ]],
  ["Thuế và lệ phí", [
    "Mỗi bên tự chịu các loại thuế phát sinh từ thu nhập của mình theo quy định.",
    "Lệ phí công chứng hợp đồng, nếu có, do hai bên chia đều.",
  ]],
  ["Bất khả kháng", [
    "Sự kiện bất khả kháng là sự kiện xảy ra khách quan, không thể lường trước và không thể khắc phục dù đã áp dụng mọi biện pháp cần thiết, gồm thiên tai, hỏa hoạn không do lỗi của các bên, dịch bệnh, chiến tranh.",
    "Bên gặp sự kiện bất khả kháng thông báo cho bên kia trong vòng 07 ngày kể từ ngày sự kiện xảy ra, kèm tài liệu chứng minh.",
    "Nếu sự kiện bất khả kháng kéo dài quá 180 ngày, mỗi bên có quyền chấm dứt hợp đồng mà không phải bồi thường.",
  ]],
  ["Thay đổi pháp luật", [
    "Nếu quy định mới làm tăng đáng kể chi phí thực hiện hợp đồng của một bên, hai bên thương lượng điều chỉnh trong vòng 60 ngày.",
    "Trong thời gian thương lượng, hai bên tiếp tục thực hiện hợp đồng theo các điều khoản hiện hành.",
  ]],
  ["Thông báo", [
    "Mọi thông báo theo hợp đồng lập bằng văn bản, gửi tới địa chỉ của các bên nêu tại phần đầu hợp đồng.",
    "Thông báo gửi bằng thư điện tử có giá trị nếu được gửi từ địa chỉ đã đăng ký và được xác nhận đã nhận.",
    "Bên thay đổi địa chỉ phải báo cho bên kia trong vòng 05 Ngày Làm Việc.",
  ]],
  ["Hiệu lực và sửa đổi", [
    "Hợp đồng có hiệu lực từ ngày ký.",
    "Mọi sửa đổi, bổ sung phải lập thành phụ lục có chữ ký của người đại diện hợp pháp của hai bên.",
    "Các phụ lục là phần không tách rời của hợp đồng; nếu phụ lục mâu thuẫn với hợp đồng thì hợp đồng được ưu tiên áp dụng.",
  ]],
  ["Luật áp dụng", [
    "Hợp đồng này được điều chỉnh và giải thích theo pháp luật Việt Nam.",
    "Nếu một điều khoản bị tuyên vô hiệu thì các điều khoản còn lại vẫn có hiệu lực.",
  ]],
  ["Điều khoản chống tham nhũng", [
    "Mỗi bên cam kết không đưa, nhận hoặc hứa hẹn bất kỳ lợi ích không chính đáng nào liên quan đến hợp đồng.",
    "Vi phạm cam kết tại Điều này là căn cứ để bên kia chấm dứt hợp đồng ngay.",
  ]],
  ["Giải quyết tranh chấp", [
    "Mọi tranh chấp trước hết được giải quyết bằng thương lượng trong vòng 30 ngày kể từ ngày một bên gửi thông báo tranh chấp.",
    "Nếu thương lượng không thành, tranh chấp được giải quyết tại Trung tâm Trọng tài Thương mại Phương Nam theo quy tắc tố tụng của trung tâm này, bởi hội đồng gồm ba trọng tài viên, tại Thành phố Hồ Chí Minh.",
    "Ngôn ngữ trọng tài là tiếng Việt; phán quyết trọng tài là chung thẩm.",
  ]],
  ["Điều khoản cuối cùng", [
    "Hợp đồng được lập thành 04 bản có giá trị như nhau, mỗi bên giữ 02 bản.",
    "Hai bên đã đọc, hiểu rõ và tự nguyện ký tên dưới đây.",
  ]],
];

// Câu đệm trung tính (không chứa số liệu) để văn bản đủ dài; gắn theo tiêu đề điều nên không lặp nguyên văn giữa các điều
const DEM = [
  (t) => `Việc thực hiện nội dung về ${t} phải phù hợp với các điều khoản khác của hợp đồng và không làm giảm quyền của bên còn lại.`,
  (t) => `Hai bên ghi nhận rằng quy định về ${t} đã được trao đổi kỹ trong quá trình đàm phán và phản ánh đúng ý chí của mỗi bên.`,
  (t) => `Khi phát sinh vướng mắc liên quan đến ${t}, bên phát hiện thông báo ngay cho bên kia để cùng xem xét, xử lý trên tinh thần hợp tác.`,
  (t) => `Mọi tài liệu, biên bản lập trong quá trình thực hiện quy định về ${t} được lưu giữ đầy đủ và xuất trình khi bên kia có yêu cầu hợp lý.`,
  (t) => `Trường hợp cần hướng dẫn chi tiết hơn về ${t}, hai bên thống nhất bằng văn bản và văn bản đó được coi là một phần của hợp đồng.`,
  (t) => `Không bên nào được viện dẫn việc bên kia chậm thực hiện quyền của mình về ${t} để coi đó là sự từ bỏ quyền.`,
  (t) => `Người đại diện của mỗi bên chịu trách nhiệm phổ biến quy định về ${t} tới các bộ phận có liên quan của mình.`,
  (t) => `Chi phí phát sinh do một bên không tuân thủ quy định về ${t} do chính bên đó gánh chịu, trừ khi hợp đồng có quy định khác.`,
  (t) => `Định kỳ hằng năm, hai bên rà soát việc thực hiện nội dung về ${t} và ghi nhận kết quả bằng biên bản.`,
  (t) => `Quy định về ${t} tiếp tục ràng buộc người kế thừa quyền, nghĩa vụ hợp pháp của mỗi bên.`,
  (t) => `Các bên hiểu rằng nội dung về ${t} được giải thích theo ngữ cảnh chung của hợp đồng và thông lệ của ngành.`,
  (t) => `Việc một bên chấp nhận ngoại lệ đối với ${t} trong một lần không tạo thành tiền lệ cho các lần sau.`,
];
const low = (s) => s[0].toLowerCase() + s.slice(1);

// 30 sửa đổi của "Đối tác": 12 cặp thay (24) + 3 thêm thuần + 2 xóa thuần + 1 đoạn thêm mới = 30
// khóa "điều.khoản" → [chữ cũ, chữ mới] (chữ cũ rỗng = thêm sau cụm neo ở phần tử thứ ba; chữ mới rỗng = xóa)
const SUA = {
  "3.2": ["09 tháng", "06 tháng"],
  "4.3": ["tăng 5%", "tăng 3%"],
  "5.1": ["10 Ngày Làm Việc", "20 Ngày Làm Việc"],
  "6.1": ["15 ngày đầu", "30 ngày đầu"],
  "6.4": ["quá 45 ngày", "quá 90 ngày"],
  "8.3": ["gấp đôi", "gấp ba"],
  "10.3": ["15 Ngày Làm Việc", "10 Ngày Làm Việc"],
  "11.3": ["07 ngày", "05 ngày"],
  "22.4": ["06 tháng", "03 tháng"],
  "23.1": ["8%", "6%"],
  "24.3": ["150%", "120%"],
  "27.3": ["180 ngày", "120 ngày"],
  "2.3": ["", ", và Bên Cho Thuê không được từ chối nếu không có lý do chính đáng", "Bên Cho Thuê"],
  "9.2": ["", " đã được thông báo cho Bên Thuê", "ban hành"],
  "16.2": ["", " và không quá hai lần mỗi quý", "trường hợp khẩn cấp"],
  "23.3": [", trừ trường hợp cố ý vi phạm", ""],
  "15.2": [" bằng văn bản", ""],
};
// 5 nhận xét của "Đối tác": khóa "điều.khoản" → [cụm được neo, nội dung]
const NHAN_XET = {
  "4.2": ["450.000 đồng", "Đề nghị giảm đơn giá thuê xuống 420.000 đồng cho mỗi mét vuông mỗi tháng."],
  "5.1": ["03 tháng Tiền Thuê", "Đề nghị giảm tiền đặt cọc còn 02 tháng Tiền Thuê."],
  "21.1": ["ngay lập tức", "Đề nghị sửa “ngay lập tức” thành “trong vòng 05 Ngày Làm Việc”."],
  "25.1": ["03 năm", "Đề nghị kéo dài thời hạn bảo mật thành 05 năm."],
  "33.1": ["30 ngày", "Đề nghị rút thời hạn thương lượng xuống 20 ngày."],
};

// 10 câu hỏi của tiêu chí 2 và số điều đúng
export const CAU_HOI = [
  ["Điều 4.2 của hợp đồng nói gì?", "4.2", "450.000"],
  ["Lãi chậm thanh toán được quy định ở điều nào và mức bao nhiêu?", "6.3", "0,05%"],
  ["Tiền đặt cọc là bao nhiêu, quy định ở điều nào?", "5.1", "03 tháng"],
  ["Điều 12.4 nói gì?", "12.4", "80 tỷ"],
  ["Thời hạn thuê là bao lâu, ở điều nào?", "3.1", "10 năm"],
  ["Quy định về bất khả kháng nằm ở điều nào? Kéo dài bao lâu thì được chấm dứt hợp đồng?", "27", "120 ngày"],
  ["Điều 18.2(b) nói gì?", "18.2(b)", "72 giờ"],
  ["Bên Thuê có được cho thuê lại không? Điều nào quy định?", "15.1", "công ty con"],
  ["Tranh chấp được giải quyết ở đâu, theo điều nào?", "33.2", "Trọng tài"],
  ["Trong Phụ lục 01, mục 2.2 nói gì?", "2.2", "30% Tiền Thuê"],
];

export function hopDong(soCauDem = 12) {
  const lists = [
    // 1: dùng cho bản thân tài liệu (không dùng)
    [`<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl>`],
    // 2: nhiều cấp gắn kiểu tiêu đề
    [
      `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading1"/><w:suff w:val="space"/><w:lvlText w:val="Điều %1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl>`,
      `<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading2"/><w:lvlText w:val="%1.%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="680" w:hanging="680"/></w:pPr></w:lvl>`,
      `<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:pStyle w:val="Heading3"/><w:lvlText w:val="%3)"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1080" w:hanging="400"/></w:pPr></w:lvl>`,
    ],
  ];
  const comments = [];
  let demAt = 0;
  /** Chữ của một khoản: áp sửa đổi của đối tác và neo nhận xét nếu có */
  const khoan = (key, text, title) => {
    let xml = "";
    let rest = text;
    const s = SUA[key];
    const c = NHAN_XET[key];
    // Vị trí cần chèn mốc trong chuỗi gốc
    const marks = [];
    if (s) {
      const anchor = s[0] || s[2];
      const at = rest.indexOf(anchor);
      if (at < 0) throw new Error(`Sửa đổi ${key}: không thấy “${anchor}”`);
      marks.push({ at: s[0] ? at : at + anchor.length, len: s[0].length, xml: (s[0] ? del(s[0]) : "") + (s[1] ? ins(s[1]) : "") });
    }
    if (c) {
      const at = rest.indexOf(c[0]);
      if (at < 0) throw new Error(`Nhận xét ${key}: không thấy “${c[0]}”`);
      const id = comments.length;
      comments.push({ text: c[1], author: "Đối tác" });
      marks.push({ at, len: c[0].length, xml: commented(id, run(c[0])) });
    }
    marks.sort((a, b) => a.at - b.at);
    let pos = 0;
    for (const m of marks) {
      if (m.at < pos) throw new Error(`Khoản ${key}: sửa đổi và nhận xét chồng nhau`);
      xml += run(rest.slice(pos, m.at)) + m.xml;
      pos = m.at + m.len;
    }
    xml += run(rest.slice(pos));
    // Câu đệm
    let more = "";
    for (let k = 0; k < soCauDem; k++) more += " " + DEM[(demAt + k) % DEM.length](low(title));
    demAt += 5;
    return xml + run(more);
  };

  const body = [
    para(run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", { b: true })),
    para(run("HỢP ĐỒNG THUÊ NHÀ XƯỞNG VÀ VĂN PHÒNG", { b: true })),
    para("Số: 07/2026/HĐT-NX (bản thử của AI Đa Nguồn, nội dung không có thật)"),
    para("Hôm nay, tại Khu công nghiệp Tân Phú, các bên gồm:"),
    para(run("BÊN CHO THUÊ: ", { b: true }) + run("Công ty Cổ phần Hạ tầng Công nghiệp Sông Lam, sau đây gọi là “Bên Cho Thuê”.")),
    para(run("BÊN THUÊ: ", { b: true }) + run("Công ty TNHH Cơ khí Chính xác Bắc Hà, sau đây gọi là “Bên Thuê”.")),
    para("Hai bên thống nhất ký hợp đồng thuê nhà xưởng và văn phòng với các điều khoản sau:"),
  ];
  DIEU.forEach(([title, clauses], a) => {
    body.push(para(title, { style: "Heading1" }));
    clauses.forEach((cl, k) => {
      const key = `${a + 1}.${k + 1}`;
      const [text, points] = Array.isArray(cl) ? cl : [cl, null];
      body.push(para(khoan(key, text, title), { style: "Heading2" }));
      (points || []).forEach((pt) => body.push(para(pt, { style: "Heading3" })));
    });
    // Một đoạn đối tác thêm mới hẳn, sau Điều 7
    if (a === 6) body.push(para(ins("Bên Cho Thuê thông báo trước ít nhất 30 ngày khi điều chỉnh phí quản lý."), { style: "Heading2", markInserted: "Đối tác" }));
  });

  // Phụ lục đánh số GÕ TAY
  const pl = (t) => para(t, { style: "Khoan" });
  body.push(
    para(run("PHỤ LỤC 01. MÔ TẢ MẶT BẰNG VÀ TIẾN ĐỘ THANH TOÁN", { b: true })),
    pl("1. Mô tả Khu Nhà Xưởng"),
    pl("1.1 Nhà xưởng số 1 có diện tích sàn 7.200 m², chiều cao thông thủy 9 mét, nền chịu tải 2 tấn trên mỗi mét vuông."),
    pl("1.2 Nhà xưởng số 2 có diện tích sàn 4.100 m², có cầu trục 5 tấn do Bên Cho Thuê lắp đặt."),
    pl("1.3 Khối văn phòng hai tầng có diện tích sàn 1.300 m², đã hoàn thiện trần, sàn và điều hòa trung tâm."),
    pl("2. Tiến độ thanh toán kỳ đầu"),
    pl("2.1 Trong vòng 10 Ngày Làm Việc kể từ ngày ký: thanh toán tiền đặt cọc."),
    pl("2.2 Trước Ngày Bàn Giao 15 ngày: thanh toán 30% Tiền Thuê của năm thuê đầu tiên."),
    pl("2.3 Phần Tiền Thuê còn lại của năm đầu được thanh toán theo quý như quy định tại Điều 6 của hợp đồng."),
    pl("a) Kỳ quý được tính theo năm dương lịch;"),
    pl("b) Kỳ lẻ đầu tiên được tính theo số ngày thực tế."),
    pl("3. Tài sản thế chấp"),
    pl("3.1 Khu Nhà Xưởng đang được thế chấp tại Ngân hàng Thương mại Cổ phần Đông Hải theo hợp đồng thế chấp đã đăng ký."),
    pl("3.2 Bên Cho Thuê cam kết ngân hàng nhận thế chấp đã chấp thuận việc cho thuê theo hợp đồng này."),
    para(""),
    para(run("ĐẠI DIỆN BÊN CHO THUÊ                                        ĐẠI DIỆN BÊN THUÊ", { b: true }))
  );
  return { body: body.join(""), lists, headingNumId: 2, comments };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const [, , out, n] = process.argv;
  if (!out) {
    console.log("Cách dùng: node kiem-thu/hop-dong-thu.mjs <tệp ra> [số câu đệm mỗi khoản]");
    process.exit(1);
  }
  await buildDocx(out, hopDong(n ? Number(n) : undefined));
  console.log(`Đã tạo ${out}`);
}
