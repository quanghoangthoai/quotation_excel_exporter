// public/js/sidebar_custom.js
$(document).ready(function() {
  console.log('▶ sidebar_custom.js loaded');

  // 1️⃣ Hàm mở cứng tất cả submenu (ERPNext v15+)
  function expandAll() {
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
  }

  // 2️⃣ Chạy lần đầu khi load Desk
  expandAll();

  // 3️⃣ Mỗi khi chuyển page/module thì bung lại
  $(document).on('page-change', function() {
    setTimeout(expandAll, 200);
  });

  // 4️⃣ (Tuỳ chọn) Giữ sidebar-toggle vẫn gập submenu được
  $('.sidebar-toggle')
    .off('click.sidebarCustom')
    .on('click.sidebarCustom', function() {
      $('.sidebar-menu .has-submenu > .collapse').slideToggle(200);
      $('.sidebar-menu .has-submenu').toggleClass('open');
      $('.dropdown-icon').toggleClass('caret-down caret-right');
    });
});
