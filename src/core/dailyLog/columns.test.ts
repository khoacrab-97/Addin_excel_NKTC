import { describe, expect, it } from "vitest";
import { detectColumns, isCompleteColumns } from "./columns";

describe("detectColumns (R1.12)", () => {
  it("nhận diện tiêu đề chuẩn", () => {
    expect(detectColumns(["STT", "Nội dung công việc", "Ngày bắt đầu", "Ngày kết thúc"])).toEqual({
      content: 1,
      start: 2,
      end: 3,
    });
  });

  it("nhận diện biến thể không dấu, viết hoa, tên khác", () => {
    expect(detectColumns(["TÊN CÔNG VIỆC", "Tu ngay", "Den ngay"])).toEqual({ content: 0, start: 1, end: 2 });
    expect(detectColumns(["Hạng mục công việc", "Thời gian bắt đầu", "Thời gian kết thúc"])).toEqual({
      content: 0,
      start: 1,
      end: 2,
    });
  });

  it("không nhầm cột nội dung phụ với cột nội dung công việc", () => {
    expect(detectColumns(["Nội dung nghiệm thu", "Nội dung công việc", "Ngày bắt đầu", "Ngày kết thúc"])).toEqual({
      content: 1,
      start: 2,
      end: 3,
    });
  });

  it("bỏ qua ô không phải text và trả về thiếu trường khi không nhận diện được", () => {
    expect(detectColumns([1, null, "Ghi chú"])).toEqual({});
  });
});

describe("isCompleteColumns", () => {
  it("yêu cầu đủ 3 cột và không trùng", () => {
    expect(isCompleteColumns({ content: 0, start: 1, end: 2 })).toBe(true);
    expect(isCompleteColumns({ content: 0, start: 1 })).toBe(false);
    expect(isCompleteColumns({ content: 0, start: 1, end: 1 })).toBe(false);
  });
});
