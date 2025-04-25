// public/js/sidebar_custom.js
$(document).ready(function() {
  // Override the original setup_sidebar method
  console.log('▶ sidebar_custom.js loaded');
  frappe.ui.Sidebar.prototype.setup_sidebar = (function(orig) {
    return function() {
      // Gọi bản gốc để render sidebar bình thường
      orig.apply(this, arguments);

      const $sb = this.$sidebar;
      
      function collapseAllSubmenus() {
        $sb.find('.sidebar-menu .has-submenu')
          .removeClass('open')
          .children('.collapse').removeClass('show')
          .hide()
          .end()
          .find('> a').attr('aria-expanded', 'false')
          .find('.dropdown-icon')
            .addClass('caret-right')
            .removeClass('caret-down');
      }

      // Ensure submenus are collapsed initially
      collapseAllSubmenus();

      // Run again after a short delay to handle dynamic updates
      setTimeout(collapseAllSubmenus, 500);

      // === BIND LẠI NÚT COLLAPSE (sidebar-toggle) ===
      $sb.find('.sidebar-toggle')
        .off('.customSidebar')
        .on('click.customSidebar', () => {
          const isCollapsed = $sb.find('.has-submenu').first().hasClass('open');
          
          $sb.find('.sidebar-menu .has-submenu').each(function() {
            const $submenu = $(this);
            const $collapse = $submenu.children('.collapse');
            const $icon = $submenu.find('> a .dropdown-icon');
            
            if (isCollapsed) {
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

      // === XỬ LÝ CLICK VÀO MENU CHA ===
      $sb.find('.sidebar-menu .has-submenu > a')
        .off('click.customSidebar')
        .on('click.customSidebar', function(e) {
          e.preventDefault();
          const $parent = $(this).parent();
          const $collapse = $parent.children('.collapse');
          const $icon = $(this).find('.dropdown-icon');
          
          $collapse.slideToggle(200);
          $parent.toggleClass('open');
          $icon.toggleClass('caret-down caret-right');
          $(this).attr('aria-expanded', function(i, attr) {
            return attr === 'true' ? 'false' : 'true';
          });
        });
    };
  })(frappe.ui.Sidebar.prototype.setup_sidebar);
});
