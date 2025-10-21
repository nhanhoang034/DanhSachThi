document.addEventListener("DOMContentLoaded", function () {
    let data = [];
    const totalCount = document.getElementById("totalCount");
    const tableBody = document.getElementById("memberTable");
    const searchInput = document.getElementById("searchInput");
    const roleFilter = document.getElementById("roleFilter");
    const examCodeInput = document.getElementById("examCode");

    async function loadCSV() {
        try {
            const response = await fetch("/static/data.csv");
            const csvData = await response.text();
            if (!csvData) { console.error("File CSV rỗng!"); return; }

            data = csvData.split(/\r?\n/).filter(line => line.trim() !== "").map(line => {
                const cells = line.split(",").map(cell => cell.trim());
                return cells;
            });

            renderTable(data);
        } catch (error) {
            console.error("Lỗi khi tải file CSV:", error);
        }
    }

    function normalize(str) {
        if (!str) return "";
        return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[đĐ]/g, "d").toLowerCase();
    }

    function renderTable(filteredData) {
        tableBody.innerHTML = "";

        filteredData.forEach((row, index) => {
            const tr = document.createElement("tr");
            
            // Tên
            const nameCell = document.createElement("td");
            nameCell.textContent = row[0];
            tr.appendChild(nameCell);

            // Mã hội viên
            const codeCell = document.createElement("td");
            codeCell.textContent = row[1];
            tr.appendChild(codeCell);

            // Quyền
            const roleCell = document.createElement("td");
            roleCell.textContent = row[2];
            tr.appendChild(roleCell);

            // Checkbox - đặt cuối cùng
            const checkboxCell = document.createElement("td");
            const checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.className = "selectMember";
            checkbox.dataset.index = index;
            checkboxCell.appendChild(checkbox);
            tr.appendChild(checkboxCell);

            tableBody.appendChild(tr);
        });

        totalCount.textContent = `Hiện có: ${filteredData.length} học viên`;
    }

    function filterAndRender() {
        const keyword = normalize(searchInput.value);
        const selectedRole = normalize(roleFilter.value);

        const filtered = data.filter(row => {
            const matchKeyword = row.some(cell => normalize(cell).includes(keyword));
            const matchRole = selectedRole === "" || normalize(row[2]) === selectedRole;
            return matchKeyword && matchRole;
        });

        renderTable(filtered);
    }

    document.getElementById("exportBtn").addEventListener("click", async () => {
        const selected = [];
        document.querySelectorAll(".selectMember:checked").forEach(cb => {
            const row = data[cb.dataset.index];
            selected.push(row[1]); // Chỉ gửi mã hội viên
        });

        if (selected.length === 0) return alert("Chưa chọn hội viên nào!");
        const examCode = examCodeInput.value.trim() || "KITHI";

        const response = await fetch("/export", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ selected: selected, exam_code: examCode })
        });

        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `DST_${examCode}.xlsx`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
    });

    searchInput.addEventListener("input", filterAndRender);
    roleFilter.addEventListener("change", filterAndRender);
    loadCSV();
});