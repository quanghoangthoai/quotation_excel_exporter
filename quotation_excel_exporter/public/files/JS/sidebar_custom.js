$(document).ready(function() {
    // Mở tất cả menu con mặc định
    $('.sidebar-menu .tree-toggle').each(function() {
        $(this).removeClass('closed').addClass('opened');
        $(this).next('.tree-children').show();
    });

    // Giữ chức năng thu gọn khi nhấp
    $('.sidebar-menu .tree-toggle').click(function() {
        $(this).toggleClass('opened closed');
        $(this).next('.tree-children').slideToggle();
    });
});