$(document).ready(function() {
  console.log('▶ sidebar_custom.js loaded');
// Check collapse elements
const collapses = document.querySelectorAll('.desk-sidebar .sidebar-item-container');
console.log('collapse count:', collapses.length, collapses);
// Check parent items with submenu
const parents = document.querySelectorAll('.desk-sidebar .sidebar-item-container');
console.log('has-submenu count:', parents.length, parents);
// Check dropdown icons
const icons = document.querySelectorAll('.desk-sidebar .sidebar-item-icon');
console.log('dropdown-icon count:', icons.length, icons);

  // 1️⃣ Function to expand all submenus (ERPNext v15+)
  function expandAll() {
    $('.desk-sidebar .sidebar-item-container')
      .addClass('show')
      .children('.sidebar-item-container')
        .addClass('show')
        .css('display', 'block')
      .end()
      .children('.sidebar-item')
        .attr('aria-expanded', 'true')
        .find('.sidebar-item-icon')
          .removeClass('sidebar-item-icon-right')
          .addClass('sidebar-item-icon-down');
  }

  // 2️⃣ Function to collapse all submenus
  function collapseAll() {
    $('.desk-sidebar .sidebar-item-container')
      .removeClass('show')
      .children('.sidebar-item-container')
        .removeClass('show')
        .css('display', 'none')
      .end()
      .children('.sidebar-item')
        .attr('aria-expanded', 'false')
        .find('.sidebar-item-icon')
          .addClass('sidebar-item-icon-right')
          .removeClass('sidebar-item-icon-down');
  }

  // 3️⃣ Run on initial Desk load: Expand all
  expandAll();

  // 4️⃣ Handle parent menu click
  $('.desk-sidebar').on('click', '.sidebar-item-container > .sidebar-item', function(e) {
    e.preventDefault();
    const $parent = $(this).parent();
    const $collapse = $parent.children('.sidebar-item-container');
    const $icon = $(this).find('.sidebar-item-icon');
    
    $collapse.slideToggle(200);
    $parent.toggleClass('show');
    $icon.toggleClass('sidebar-item-icon-down sidebar-item-icon-right');
    $(this).attr('aria-expanded', $parent.hasClass('show'));
  });

  // 5️⃣ Handle collapse button
  $('.sidebar-toggle').on('click', function() {
    const $submenus = $('.desk-sidebar .sidebar-item-container');
    const isAnyOpen = $submenus.filter('.show').length > 0;
    
    $submenus.each(function() {
      const $submenu = $(this);
      const $collapse = $submenu.children('.sidebar-item-container');
      const $icon = $submenu.find('> .sidebar-item .sidebar-item-icon');
      
      if (isAnyOpen) {
        $submenu.removeClass('show');
        $collapse.slideUp(200);
        $icon.removeClass('sidebar-item-icon-down').addClass('sidebar-item-icon-right');
        $submenu.find('> .sidebar-item').attr('aria-expanded', 'false');
      } else {
        $submenu.addClass('show');
        $collapse.slideDown(200);
        $icon.removeClass('sidebar-item-icon-right').addClass('sidebar-item-icon-down');
        $submenu.find('> .sidebar-item').attr('aria-expanded', 'true');
      }
    });

    // Trigger a window resize event after the animation completes
    setTimeout(() => {
      $(window).trigger('resize');
    }, 200);
  });

  // 6️⃣ Handle page-change
  $(document).on('page-change', function() {
    setTimeout(() => {
      expandAll();
      $(window).trigger('resize'); // Trigger resize to ensure charts update
    }, 200);
  });

  // 7️⃣ Debounce resize events to prevent chart redraw issues
  let resizeTimeout;
  $(window).on('resize', function() {
    clearTimeout(resizeTimeout);
    resizeTimeout = setTimeout(() => {
      // Ensure charts are updated only if their container is visible
      $('.chart-container canvas').each(function() {
        const canvas = $(this)[0];
        if (canvas.offsetParent !== null && canvas.chart) {
          canvas.chart.resize();
        }
      });
    }, 100);
  });
});