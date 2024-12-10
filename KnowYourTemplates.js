function showTable(tabId) { 
    const sections = document.querySelectorAll('.styled-table'); 
    sections.forEach(section => { 
        section.style.display = 'none'; 
    }); 
    const selectedSection = document.getElementById(tabId); 
    if (selectedSection) { 
        selectedSection.style.display = 'table'; 
    } 
    const tabs = document.querySelectorAll('.tableIndex p'); 
    tabs.forEach(tab => { tab.classList.remove('selected'); 

    }); 
    const selectedTab = document.getElementById(tabId + '_Tab'); 
    if (selectedTab) { 
        selectedTab.classList.add('selected'); 
    }
} 
document.getElementById('CASS_Tab').addEventListener('click', function() { 
    showTable('CASS'); 
}); 
document.getElementById('NORG_Tab').addEventListener('click', function() { 
    showTable('NORG'); 
}); 
document.getElementById('CLOS_Tab').addEventListener('click', function() { 
    showTable('CLOS'); 
}); 

showTable('CASS'); // Show Waterfall table by default