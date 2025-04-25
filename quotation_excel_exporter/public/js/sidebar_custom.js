// public/js/sidebar_custom.js
frappe.ui.Sidebar.prototype.setup_sidebar = (function(orig) {
  return function() {
    // Gọi bản gốc để render sidebar bình thường
    orig.apply(this, arguments);

    const $sb = this.$sidebar;

    // === THIẾT LẬP TRẠNG THÁI BAN ĐẦU - ẨN TẤT CẢ SUBMENU ===
    $sb.find('.sidebar-menu .has-submenu')
       .removeClass('open')
       .children('.collapse').removeClass('show')
       .hide()
       .siblings('a').attr('aria-expanded', 'false')
       .find('.dropdown-icon')
         .addClass('caret-right')
         .removeClass('caret-down');

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
      $(this).attr('aria-expanded', function(i, attr) {
        return attr === 'true' ? 'false' : 'true';
      });
    });
  };
})(frappe.ui.Sidebar.prototype.setup_sidebar);
