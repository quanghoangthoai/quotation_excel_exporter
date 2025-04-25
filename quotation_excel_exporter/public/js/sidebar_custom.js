// public/js/sidebar_custom.js
frappe.ready(() => {
  console.log('▶ sidebar_custom.js loaded');

  // Hàm mở tất cả submenu (ERPNext v15+)
  const expandAll = () => {
    $('.sidebar-menu .has-submenu')
      .addClass('open')
      .children('.collapse')
        .addClass('show')
        .css('display', 'block')
      .end()
      .children('a')
        .attr('aria-expanded', 'true')
        .find('.dropdown-icon')
          .removeClass('caret-right')
          .addClass('caret-down');
  };

  // Chạy ngay lần đầu
  expandAll();

  // Mỗi khi chuyển module/page
  $(document).on('page-change', () => {
    setTimeout(expandAll, 200);
  });

  // (Tuỳ chọn) Giữ nút collapse sidebar vẫn gập submenu được
  $('.sidebar-toggle').off('click.sidebarCustom').on('click.sidebarCustom', () => {
    $('.sidebar-menu .has-submenu > .collapse')
      .slideToggle(200);
    $('.sidebar-menu .has-submenu')
      .toggleClass('open');
    $('.dropdown-icon')
      .toggleClass('caret-down caret-right');
  });
});
