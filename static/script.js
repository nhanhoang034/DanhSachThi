async function exportFile() {
  const selected = Array.from(document.querySelectorAll('input[type=checkbox]:checked'))
    .map(cb => cb.value);
  const examCode = document.getElementById('exam_code').value.trim();

  if (!selected.length || !examCode) {
    alert("Vui lòng chọn hội viên và nhập mã kỳ thi!");
    return;
  }

  try {
    const response = await axios.post('/export', {
      selected: selected,
      exam_code: examCode
    }, { responseType: 'blob' });

    const url = window.URL.createObjectURL(new Blob([response.data]));
    const a = document.createElement('a');
    a.href = url;
    a.download = `DST_${examCode}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
  } catch (err) {
    alert("Lỗi khi xuất file!");
    console.error(err);
  }
}
