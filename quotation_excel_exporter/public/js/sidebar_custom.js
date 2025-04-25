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
  };
})(frappe.ui.Sidebar.prototype.setup_sidebar);
