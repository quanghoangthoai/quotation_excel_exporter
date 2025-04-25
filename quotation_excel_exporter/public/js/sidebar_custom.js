// Monkey-patch lại phương thức setup_sidebar
frappe.ui.Sidebar.prototype.setup_sidebar = (function(orig) {
  return function() {
    // gọi bản gốc để dựng sidebar
    orig.apply(this, arguments);

    const $sb = this.$sidebar;

    // MỞ TẤT CẢ menu con ngay lập tức
    $sb.find('.side-nav-item-with-child')
       .addClass('open')
       .find('> .indicator')
       .text('▾');
    $sb.find('.nested').show();

    // BIND lại nút collapse sau khi đã remove cũ
    $sb.find('.sidebar-toggle')
       .off('click.customSidebar')           // tránh chồng event
       .on('click.customSidebar', () => {
         // ẩn/hiện submenu
         $sb.find('.nested').slideToggle(200);
         // cập nhật class open
         $sb.find('.side-nav-item-with-child').toggleClass('open');

         // và đổi icon tương ứng
         $sb.find('.side-nav-item-with-child > .indicator').each(function() {
           const $i = $(this);
           $i.text($i.closest('.side-nav-item-with-child').hasClass('open') ? '▾' : '▸');
         });
       });
  };
})(frappe.ui.Sidebar.prototype.setup_sidebar);
