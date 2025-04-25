// public/js/sidebar_custom.js
frappe.ui.Sidebar.prototype.setup_sidebar = (function(orig) {
  return function() {
    // Gọi bản gốc để render sidebar bình thường
    orig.apply(this, arguments);

    const $sb = this.$sidebar;

    // === MỞ CỨNG TẤT CẢ MENU CON (v15+) ===
    $sb.find('.sidebar-menu .has-submenu')
       .addClass('open')
       .children('.collapse').addClass('show')
       .siblings('a').attr('aria-expanded', 'true')
       .find('.dropdown-icon')
         .removeClass('caret-right')
         .addClass('caret-down');

    // === BIND LẠI NÚT COLLAPSE (sidebar-toggle) ===
    $sb.find('.sidebar-toggle')
       .off('.customSidebar')
       .on('click.customSidebar', () => {
         // Toggle phần collapse của từng submenu
         $sb.find('.sidebar-menu .has-submenu > .collapse')
            .slideToggle(200);
         // Toggle class open để icon & style cập nhật
         $sb.find('.sidebar-menu .has-submenu')
            .toggleClass('open');
         // Đổi icon mũi tên
         $sb.find('.dropdown-icon')
            .toggleClass('caret-down caret-right');
       });

    // === XỬ LÝ CLICK VÀO MENU CHA ===
    $sb.find('.sidebar-menu .has-submenu > a').on('click', function(e) {
      e.preventDefault();
      const $parent = $(this).parent();
      const $collapse = $parent.children('.collapse');
      
      // Toggle submenu
      $collapse.slideToggle(200);
      $parent.toggleClass('open');
      
      // Update icon
      $(this).find('.dropdown-icon')
        .toggleClass('caret-down caret-right');
    });

    // === XỬ LÝ TRẠNG THÁI BAN ĐẦU ===
    // Đảm bảo tất cả submenu đều hiển thị khi mới load
    $sb.find('.sidebar-menu .has-submenu > .collapse').show();
    $sb.find('.sidebar-menu .has-submenu').addClass('open');
    $sb.find('.dropdown-icon').addClass('caret-down').removeClass('caret-right');
  };
})(frappe.ui.Sidebar.prototype.setup_sidebar);
