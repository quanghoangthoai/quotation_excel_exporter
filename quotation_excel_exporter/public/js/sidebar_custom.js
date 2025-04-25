$(document).ready(function() {
  console.log('▶ sidebar_custom.js loaded');

  // 1️⃣ Hàm mở tất cả submenu (ERPNext v15+)
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

  // 2️⃣ Hàm đóng tất cả submenu
  function collapseAll() {
    $('.sidebar-menu .has-submenu')
      .removeClass('open')
      .children('.collapse')
        .removeClass('show')
        .css('display', 'none')
      .end()
      .children('a')
        .attr('aria-expanded', 'false')
        .find('.dropdown-icon')
          .addClass('caret-right')
          .removeClass('caret-down');
  }

  // 3️⃣ Chạy lần đầu khi load Desk: Mở tất cả
  expandAll();

  // 4️⃣ Xử lý click vào menu cha
  $('.sidebar-menu').on('click', '.has-submenu > a', function(e) {
    e.preventDefault();
    const $parent = $(this).parent();
    const $collapse = $parent.children('.collapse');
    const $icon = $(this).find('.dropdown-icon');
    
    $collapse.slideToggle(200);
    $parent.toggleClass('open');
    $icon.toggleClass('caret-down caret-right');
    $(this).attr('aria-expanded', $parent.hasClass('open'));
  });

  // 5️⃣ Xử lý nút collapse
  $('.sidebar-toggle').on('click', function() {
    const $submenus = $('.sidebar-menu .has-submenu');
    const isAnyOpen = $submenus.filter('.open').length > 0;
    
    $submenus.each(function() {
      const $submenu = $(this);
      const $collapse = $submenu.children('.collapse');
      const $icon = $submenu.find('> a .dropdown-icon');
      
      if (isAnyOpen) {
        $submenu.removeClass('open');
        $collapse.slideUp(200);
        $icon.removeClass('caret-down').addClass('caret-right');
        $submenu.find('> a').attr('aria-expanded', 'false');
      } else {
        $submenu.addClass('open');
        $collapse.slideDown(200);
        $icon.removeClass('caret-right').addClass('caret-down');
        $submenu.find('> a').attr('aria-expanded', 'true');
      }
    });
  });

  // 6️⃣ (Tùy chọn) Nếu muốn áp dụng lại trạng thái mặc định (mở tất cả) sau khi chuyển trang
  $(document).on('page-change', function() {
    setTimeout(expandAll, 200);
  });
});