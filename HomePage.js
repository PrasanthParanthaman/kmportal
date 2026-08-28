var angleStart = -360;

// jquery rotate animation
function rotate(li,d) {
    $({d:angleStart}).animate({d:d}, {
        step: function(now) {
            $(li)
               .css({ transform: 'rotate('+now+'deg)' })
               .find('label')
                  .css({ transform: 'rotate('+(-now)+'deg)' });
        }, duration: 0
    });
}

// show / hide the options
function toggleOptions(s) {
    $(s).toggleClass('open');
    var li = $(s).find('li');
    var angles = [-5, 31, 67, 103, 139, 175]; // Your specific angles
    for (var i = 0; i < li.length; i++) {
        var d = angles[i];
        $(s).hasClass('open') ? rotate(li[i], d) : rotate(li[i], angleStart);
    }
}


$('.selector button').click(function(e) {
    toggleOptions($(this).parent());
});

setTimeout(function() { toggleOptions('.selector'); }, 100);



  // ----- On render -----
  $(document).ready(function() {
    // var mode = "home";
  
    // Event delegation for main container click
    
  
    // ----- On render -----
    makeRadial({
        el: $("#radial"),
        // el: $(".selector"),
        radials: 100
    });
  
    function makeRadial(options) {
        if (options && options.el) {
            var el = options.el;
            var radials = 60;
            if (options.radials) {
                radials = options.radials;
            }
            var degrees = 360 / radials;
            var i = 0;
            for (i = 0; i < radials / 2; i++) {
                var newTick = $('<div class="tick"></div>')
                    .css({
                        "-moz-transform": "rotate(" + i * degrees + "deg)"
                    })
                    .css({
                        "-webkit-transform": "rotate(" + i * degrees + "deg)"
                    })
                    .css({
                        transform: "rotate(" + i * degrees + "deg)"
                    });
                el.prepend(newTick);
            }
        }
    }
  });


  
//For Side Tab Animation and Blinking Stars
document.addEventListener('DOMContentLoaded', function () {
    //Side Tab Animation 
    var tabs = [
        { btnId: 'bootcampBtn', tabClass: 'tab-bootCamp' },
        { btnId: 'driveinBtn', tabClass: 'tab-driveIn' },
        { btnId: 'dashboardBtn', tabClass: 'tab-dashboard' },
        { btnId: 'automationBtn', tabClass: 'tab-automation' },
        { btnId: 'fameBtn', tabClass: 'tab-fame' },
        { btnId: 'fuelupBtn', tabClass: 'tab-fuelup' }
    ];

    tabs.forEach(function(tab) {
        var btn = document.getElementById(tab.btnId);
        var tabElement = document.querySelector('.' + tab.tabClass);

        btn.addEventListener('mouseenter', function () {
            tabElement.style.display = 'flex'; // Ensure the tab is displayed
            tabElement.classList.remove('hide');
            tabElement.classList.add('show');
        });

        btn.addEventListener('mouseleave', function () {
            tabElement.classList.remove('show');
            tabElement.classList.add('hide');
            setTimeout(function() {
                tabElement.style.display = 'none';
            }, 500); 
        });
    });
});

// // --- Profile ---

document.addEventListener('DOMContentLoaded', function () {
    // Profile Menu Logic
    var profileIcon = document.getElementById('profileIcon');
    var profileMenu = document.querySelector('.profile-dropdown .profile-menu');
    var blurredBackground = document.getElementById('blurredBackground');
    var messageBox = document.getElementById('messageBox');
    
    // Quick Links Dropdown Logic
    var quickLinks = document.getElementById('quickLinks');
    var quickLinksMenu = document.getElementById('quickLinksMenu'); // Added ID
    
    // Toggle the Quick Links menu on click
    quickLinks.addEventListener('click', function (event) {
    event.preventDefault();
    
    quickLinksMenu.style.display = quickLinksMenu.style.display === 'block' ? 'none' : 'block';
    });
    
    // Toggle the profile menu on profile picture click
    profileIcon.addEventListener('click', function (event) {
    event.preventDefault();
    console.log('Profile icon clicked'); // Add this line
    console.log('Current display style:', profileMenu.style.display);
    profileMenu.style.display = profileMenu.style.display === 'block' ? 'none' : 'block';
    console.log('New display style:', profileMenu.style.display);
});
    
    // Close the profile and quick links menu if clicking outside
    document.addEventListener('click', function (event) {
    var isClickInsideProfile = profileIcon.contains(event.target) || profileMenu.contains(event.target);
    var isClickInsideQuickLinks = quickLinks.contains(event.target) || quickLinksMenu.contains(event.target);
    
    if (!isClickInsideProfile) {
        profileMenu.style.display = 'none';
    }
    
    if (!isClickInsideQuickLinks) {
        quickLinksMenu.style.display = 'none';
    }
    });
    
    // Show logout confirmation dialog
    window.showLogoutConfirmation = function () {
    blurredBackground.style.display = 'block';
    messageBox.style.display = 'block';
    profileMenu.style.display = 'none';
    };
    
    // Confirm logout
    window.confirmLogout = function () {
    window.location.href = "/Logout"; // Redirect to actual logout URL
    };
    
    // Cancel logout
    window.cancelLogout = function () {
    blurredBackground.style.display = 'none';
    messageBox.style.display = 'none';
    };
    
    // Retrieve the stored name and email
    const name = localStorage.getItem('userName');
    const email = localStorage.getItem('userEmail');
    
    // Display the name and email in the profile section
    if (name && email) {
    document.querySelector('.profile-dropdown .profile-menu p').textContent = name;
    document.querySelector('.profile-dropdown .profile-menu p:nth-of-type(2)').textContent = email + "@cognizant.com";
    }
    });

// Pages object with URLs and titles for suggestions
const pages = {
    'bootcamp': { url: '/Bootcamp', title: 'Boot Camp' },
    'drivein': { url: '/DriveInDomain', title: 'Drive In' },
    'dashboard': { url: '/Dashboard', title: 'Dashboard' },
    'automation': { url: '/Automation', title: 'Automation' },
    'fuelup': { url: '/FuelUP', title: 'Fuel Up' },
    'selfstart': { url: '/Bootcamp', title: 'Self Start' },
    'autoknowall': { url: '/Bootcamp', title: 'Auto Know All' },
    'applicationmanual': { url: '/Bootcamp', title: 'Application Manual' },
    'elearning': { url: '/Bootcamp', title: 'E Learning' },
    'guidelines': { url: '/Bootcamp', title: 'Guidelines' },
    'toolkit': { url: '/Bootcamp', title: 'Toolkit' },
    'templates': { url: '/Bootcamp', title: 'Templates' },
    'assesment': { url: '/Bootcamp', title: 'Assessment' },
    'knowyourdomain': { url: '/Knowyourdomain', title: 'Know Your Domain' },
    'knowyourcustomer': { url: '/Kyc', title: 'Know Your Customer' },
    'knowyourapplication': { url: '/Kya', title: 'Know Your Application' },
    'termstoknow': { url: '/TermsToKnow', title: 'Terms To Know' },
    'technologies': { url: '/Technology', title: 'Technologies' },
    'casestudy': { url: '/Casestudies', title: 'Case Study' },
    'newsletter': { url: '/NewsLetter', title: 'Newsletter' },
    'domainvideos': { url: '/DomainVideo', title: 'Domain Videos' },
    'specializedautomation': { url: '/SpecializedAutomation', title: 'Specialized Automation' },
    'keyaccelerator': { url: '/KeyAccelarator', title: 'Key Accelerator' },
    'bluebolt': { url: '/Automation', title: 'Blue Bolt' },
    'projectdetails': { url: '/Dashboard', title: 'Project Details' },
    'sme': { url: '/KeySubjectMatterExpert', title: 'SME' },
    'subjectmatterexpert': { url: '/SubjectMatterExpert', title: 'Subject Matter Expert' },
    'transformationplan': { url: '/TransformationPlan', title: 'Transformation Plan' },
    'accountzone': { url: '/Dashboard', title: 'Account Zone' },
    'tools': { url: '/Dashboard', title: 'Tools' },
    'skillmatrix': { url: '/Dashboard', title: 'Skill Matrix' }
};

// Helper function to capitalize the first letter and add spaces between words
function formatKeyword(keyword) {
    return keyword.replace(/([A-Z])/g, ' $1') // Add space before each uppercase letter
                  .replace(/^./, str => str.toUpperCase()) // Capitalize the first letter
                  .trim(); // Remove any leading spaces
}

document.addEventListener('DOMContentLoaded', () => {
    // Add your JavaScript code here
    const searchInput = document.getElementById('searchInput');
    const suggestions = document.getElementById('suggestions');
    let selectedIndex = -1;

    // Ensure searchInput and suggestions exist
    if (!searchInput || !suggestions) {
        console.error('Search input or suggestions container not found in the DOM.');
        return;
    }

    // Function to suggest search keywords
    function suggestSearch(query) {
        suggestions.innerHTML = '';
        suggestions.style.display = 'none';

        if (query.length === 0) {
            return; // Exit if the query is empty
        }

        // Filter keywords based on the query
        const matches = Object.keys(pages).filter(page => page.startsWith(query.toLowerCase()));

        if (matches.length > 0) {
            suggestions.style.display = 'block'; // Show suggestions container
            matches.forEach(match => {
                const suggestionDiv = document.createElement('div');
                suggestionDiv.textContent = pages[match].title; // Use the title property for display
                suggestionDiv.classList.add('suggestion-item');
                suggestionDiv.onclick = () => {
                    searchInput.value = pages[match].title; // Set input value to the title
                    suggestions.style.display = 'none'; // Hide suggestions
                    performSearch(match); // Perform search
                };
                suggestions.appendChild(suggestionDiv); // Add suggestion to the container
            });
        }
    }

    // Function to navigate to matched page
    function performSearch(key) {
        const query = key || searchInput.value.toLowerCase().replace(/\s+/g, ''); // Remove spaces for matching
        const match = Object.keys(pages).find(page => query.includes(page));

        if (match) {
            window.location.href = pages[match].url; // Use the URL property for navigation
        } else {
            alert('No matching page found. Please try again.');
        }
    }

    // Add event listeners for input and keyboard navigation
    searchInput.addEventListener('input', () => {
        selectedIndex = -1; // Reset selection index when input changes
        suggestSearch(searchInput.value);
    });

    searchInput.addEventListener('keydown', (e) => {
        const items = suggestions.getElementsByClassName('suggestion-item');
        if (e.key === 'ArrowDown') {
            e.preventDefault(); // Prevent cursor movement in input
            if (items.length > 0) {
                selectedIndex = (selectedIndex + 1) % items.length;
                updateSelection(items);
            }
        } else if (e.key === 'ArrowUp') {
            e.preventDefault(); // Prevent cursor movement in input
            if (items.length > 0) {
                selectedIndex = (selectedIndex - 1 + items.length) % items.length;
                updateSelection(items);
            }
        } else if (e.key === 'Enter') {
            e.preventDefault(); // Prevent form submission
            if (selectedIndex >= 0 && items[selectedIndex]) {
                searchInput.value = items[selectedIndex].textContent;
                suggestions.style.display = 'none';
                performSearch();
            }
        }
    });

    // Hide suggestions when clicking outside the search bar
    document.addEventListener('click', (e) => {
        if (!searchInput.contains(e.target) && !suggestions.contains(e.target)) {
            suggestions.style.display = 'none';
        }
    });

    // Function to update the selected suggestion
    function updateSelection(items) {
        for (let i = 0; i < items.length; i++) {
            items[i].classList.remove('selected');
        }
        if (selectedIndex >= 0) {
            items[selectedIndex].classList.add('selected');
            items[selectedIndex].scrollIntoView({ block: 'nearest' }); // Ensure the selected item is visible
        }
    }
});

document.addEventListener('DOMContentLoaded', () => {
        // Fetch the last update time
        fetch('/last-update')
            .then(response => response.json())
            .then(data => {
                // Safely access the element
                const lastUpdateDiv = document.getElementById('last-update');
                const updateTimeSpan = document.getElementById('update-time');

                // Check if the elements exist before proceeding
                if (lastUpdateDiv && updateTimeSpan) {
                    // Update the time in the span
                    updateTimeSpan.textContent = data.last_update_time;

                    // Display the entire "Last Update" div
                    lastUpdateDiv.style.display = 'block';

                    // Hide the entire div after 5 seconds
                    setTimeout(() => {
                        lastUpdateDiv.style.display = 'none'; // Hides both "Last Update:" and the time
                    }, 5000);
                } else {
                    console.error('Error: Element not found in the DOM.');
                }
            })
            .catch(error => {
                console.error('Error fetching last update time:', error);
            });
    });










