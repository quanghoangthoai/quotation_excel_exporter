// public/js/sidebar_custom.js
frappe.provide('quotation_excel_exporter');

quotation_excel_exporter.setup_sidebar = function() {
  if (!frappe.ui.Sidebar || !frappe.ui.Sidebar.prototype) {
    // If Sidebar is not ready, try again in 100ms
    setTimeout(quotation_excel_exporter.setup_sidebar, 100);
    return;
  }

  // Only patch if not already patched
  if (frappe.ui.Sidebar.prototype._patched) return;
  frappe.ui.Sidebar.prototype._patched = true;

  const originalSetupSidebar = frappe.ui.Sidebar.prototype.setup_sidebar;
  
  frappe.ui.Sidebar.prototype.setup_sidebar = function() {
    try {
      // Call original method
      originalSetupSidebar.apply(this, arguments);

      const $sb = this.$sidebar;
      if (!$sb || !$sb.length) return;
      
      const collapseAllSubmenus = () => {
        try {
          const $submenus = $sb.find('.sidebar-menu .has-submenu');
          if (!$submenus.length) return;

          $submenus.each(function() {
            const $item = $(this);
            const $collapse = $item.children('.collapse');
            const $link = $item.find('> a');
            const $icon = $link.find('.dropdown-icon');

            $item.removeClass('open');
            $collapse.removeClass('show').hide();
            $link.attr('aria-expanded', 'false');
            $icon.addClass('caret-right').removeClass('caret-down');
          });
        } catch (err) {
          console.warn('Error in collapseAllSubmenus:', err);
        }
      };

      // Initial collapse with retry
      const initializeCollapse = () => {
        collapseAllSubmenus();
        // Retry after DOM updates
        setTimeout(collapseAllSubmenus, 500);
      };
      
      initializeCollapse();

      // Handle collapse button
      $sb.find('.sidebar-toggle')
        .off('.customSidebar')
        .on('click.customSidebar', () => {
          try {
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
          } catch (err) {
            console.warn('Error in collapse button handler:', err);
          }
        });

      // Handle submenu clicks
      $sb.find('.sidebar-menu .has-submenu > a')
        .off('.customSidebar')
        .on('click.customSidebar', function(e) {
          try {
            e.preventDefault();
            const $link = $(this);
            const $parent = $link.parent();
            const $collapse = $parent.children('.collapse');
            const $icon = $link.find('.dropdown-icon');
            
            $collapse.slideToggle(200);
            $parent.toggleClass('open');
            $icon.toggleClass('caret-down caret-right');
            $link.attr('aria-expanded', function(i, attr) {
              return attr === 'true' ? 'false' : 'true';
            });
          } catch (err) {
            console.warn('Error in submenu click handler:', err);
          }
        });
    } catch (err) {
      console.warn('Error in setup_sidebar:', err);
    }
  };
};

// Initialize when document is ready
$(document).ready(function() {
  console.log('▶ sidebar_custom.js loaded');
  quotation_excel_exporter.setup_sidebar();
});

// Also try to initialize when Frappe is ready
$(document).on('frappe.ready', function() {
  console.log('▶ Frappe ready - initializing sidebar');
  quotation_excel_exporter.setup_sidebar();
});
