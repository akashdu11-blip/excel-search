let allProducts = [];
let fuse = null;
let activeIndex = -1;
let currentResults = [];

const searchInput = document.getElementById('searchInput');
const dropdown = document.getElementById('dropdown');
const noResultsEl = document.getElementById('noResults');
const totalRecordsEl = document.getElementById('totalRecords');
const loadingSpinner = document.getElementById('loadingSpinner');
const uploadForm = document.getElementById('uploadForm');
const uploadStatus = document.getElementById('uploadStatus');

function debounce(fn, delay) {
  let timer = null;
  return function (...args) {
    if (timer) clearTimeout(timer);
    timer = setTimeout(() => fn.apply(this, args), delay);
  };
}

async function fetchProducts() {
  loadingSpinner.classList.remove('hidden');
  try {
    const res = await fetch('/products');
    if (!res.ok) throw new Error('Failed to load products');
    const data = await res.json();
    allProducts = data.products || [];
    totalRecordsEl.textContent = `Total: ${allProducts.length}`;

    fuse = new Fuse(allProducts, {
      keys: ['product_no', 'description'],
      threshold: 0.3,
      ignoreLocation: true,
      includeMatches: true,
      minMatchCharLength: 1
    });
  } catch (err) {
    console.error(err);
  } finally {
    loadingSpinner.classList.add('hidden');
  }
}

function clearDropdown() {
  dropdown.innerHTML = '';
  dropdown.classList.add('hidden');
  noResultsEl.classList.add('hidden');
  activeIndex = -1;
  currentResults = [];
}

function highlightMatches(text, matchesForField) {
  if (!matchesForField || !matchesForField.indices || matchesForField.indices.length === 0) {
    return text;
  }

  let result = '';
  let lastIndex = 0;

  matchesForField.indices.forEach(([start, end]) => {
    if (start > text.length || end > text.length) return;
    result += text.slice(lastIndex, start);
    result += `<span class="highlight">${text.slice(start, end + 1)}</span>`;
    lastIndex = end + 1;
  });

  result += text.slice(lastIndex);
  return result;
}

function renderDropdown(results) {
  dropdown.innerHTML = '';

  const header = document.createElement('div');
  header.className = 'dropdown-header';
  header.innerHTML = `
    <span>Product No</span>
    <span>Description</span>
    <span>Rate</span>
  `;
  dropdown.appendChild(header);

  results.forEach((res, index) => {
    const item = res.item;
    const matches = res.matches || [];

    const productMatch = matches.find((m) => m.key === 'product_no');
    const descriptionMatch = matches.find((m) => m.key === 'description');

    const productHtml = highlightMatches(item.product_no || '', productMatch);
    const descriptionHtml = highlightMatches(item.description || '', descriptionMatch);
    const rateText = item.rate != null ? item.rate : '';

    const row = document.createElement('div');
    row.className = 'dropdown-item';
    row.dataset.index = index;
    row.innerHTML = `
      <span>${productHtml}</span>
      <span>${descriptionHtml}</span>
      <span>${rateText}</span>
    `;

    row.addEventListener('mousedown', (e) => {
      e.preventDefault();
      selectItem(index);
    });

    dropdown.appendChild(row);
  });

  dropdown.classList.remove('hidden');
}

function selectItem(index) {
  if (index < 0 || index >= currentResults.length) return;
  const item = currentResults[index].item;
  searchInput.value = item.description || '';
  clearDropdown();
}

function updateActiveItem(nextIndex) {
  const items = dropdown.querySelectorAll('.dropdown-item');
  if (!items.length) return;

  if (activeIndex >= 0 && activeIndex < items.length) {
    items[activeIndex].classList.remove('active');
  }

  activeIndex = nextIndex;

  if (activeIndex >= 0 && activeIndex < items.length) {
    const activeEl = items[activeIndex];
    activeEl.classList.add('active');
    const parent = dropdown;
    const offsetTop = activeEl.offsetTop;
    const scrollBottom = parent.scrollTop + parent.clientHeight;
    if (offsetTop < parent.scrollTop) {
      parent.scrollTop = offsetTop;
    } else if (offsetTop + activeEl.clientHeight > scrollBottom) {
      parent.scrollTop = offsetTop + activeEl.clientHeight - parent.clientHeight;
    }
  }
}

const handleSearchInput = debounce(() => {
  const query = searchInput.value.trim();
  if (!query || !fuse) {
    clearDropdown();
    return;
  }

  const results = fuse.search(query, { limit: 20 });
  currentResults = results;

  if (!results.length) {
    dropdown.innerHTML = '';
    dropdown.classList.add('hidden');
    noResultsEl.classList.remove('hidden');
    activeIndex = -1;
    return;
  }

  noResultsEl.classList.add('hidden');
  renderDropdown(results);
  activeIndex = -1;
}, 300);

searchInput.addEventListener('input', handleSearchInput);

searchInput.addEventListener('keydown', (e) => {
  const itemsCount = currentResults.length;

  if (e.key === 'ArrowDown') {
    if (!itemsCount) return;
    e.preventDefault();
    const next = activeIndex < itemsCount - 1 ? activeIndex + 1 : 0;
    updateActiveItem(next);
  } else if (e.key === 'ArrowUp') {
    if (!itemsCount) return;
    e.preventDefault();
    const next = activeIndex > 0 ? activeIndex - 1 : itemsCount - 1;
    updateActiveItem(next);
  } else if (e.key === 'Enter') {
    if (!itemsCount) return;
    e.preventDefault();
    const indexToSelect = activeIndex >= 0 ? activeIndex : 0;
    selectItem(indexToSelect);
  } else if (e.key === 'Escape') {
    clearDropdown();
  }
});

document.addEventListener('click', (e) => {
  if (!dropdown.contains(e.target) && e.target !== searchInput) {
    clearDropdown();
  }
});

uploadForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  const filesInput = document.getElementById('excelFiles');
  const files = filesInput.files;
  if (!files || !files.length) {
    uploadStatus.textContent = 'Please select at least one Excel file.';
    uploadStatus.className = 'status-text error';
    return;
  }

  const formData = new FormData();
  for (let i = 0; i < files.length; i++) {
    formData.append('files', files[i]);
  }

  uploadStatus.textContent = 'Uploading and processing...';
  uploadStatus.className = 'status-text';

  try {
    const res = await fetch('/upload', {
      method: 'POST',
      body: formData
    });

    const data = await res.json();
    if (!res.ok) {
      uploadStatus.textContent = data.error || 'Upload failed';
      uploadStatus.className = 'status-text error';
      return;
    }

    const added = data.added != null ? data.added : data.total_records;
    uploadStatus.textContent = added === data.total_records
      ? `Processed successfully. Total records: ${data.total_records}`
      : `Added ${added} records. Total: ${data.total_records}`;
    uploadStatus.className = 'status-text success';

    await fetchProducts();
    clearDropdown();
    searchInput.value = '';
  } catch (err) {
    console.error(err);
    uploadStatus.textContent = 'Failed to upload or process files.';
    uploadStatus.className = 'status-text error';
  }
});

window.addEventListener('DOMContentLoaded', () => {
  fetchProducts();
});

