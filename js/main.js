//Global Variables
let heads = [];
let Categories1Sheet1 = [];
let Categories1Sheet2 = [];
let Categories1Sheet3 = [];
let Categories1Sheet4 = [];
let Categories1Sheet5 = [];
let Categories1Sheet6 = [];
let Categories1Sheet7 = [];
let arr1sheet1 = [];
let arr2sheet1 = [];
let arr1sheet2 = [];
let arr1sheet3 = [];
let arr1sheet4 = [];
let arr1sheet5 = [];
let arr2sheet5 = [];
let arr1sheet6 = [];
let arr2sheet6 = [];
let arr1sheet7 = [];
let arr2sheet7 = [];
let sumTotalATM = 0;
let chartVisableCount = 60;
let contaniers = document.querySelectorAll('[id^="container"]');
let container1 = document.querySelector("#container-1.active");

//Animation Function

Math.easeOutBounce = (pos) => {
  if (pos < 1 / 2.75) {
    return 7.5625 * pos * pos;
  } else if (pos < 2 / 2.75) {
    return 7.5625 * (pos -= 1.5 / 2.75) * pos + 0.75;
  } else if (pos < 2.5 / 2.75) {
    return 7.5625 * (pos -= 2.25 / 2.75) * pos + 0.9375;
  } else {
    return 7.5625 * (pos -= 2.625 / 2.75) * pos + 0.984375;
  }
};

//Start Uploading files
document.getElementById("btn-submit").addEventListener("click", () => {
  let upload = document.getElementById("file");
  //Sheet 1 importing
  readXlsxFile(upload.files[0], { sheet: "installed VS removed" }).then(
    (data) => {
      heads.push(data[0]);
      data.forEach((r) => {
        if (r[0] != "Customer") {
          Categories1Sheet1.push(r[0]);
          arr1sheet1.push(r[2]);
          arr2sheet1.push(r[3]);
        }
      });
    }
  );
  //Sheet 2 importing
  readXlsxFile(upload.files[0], { sheet: "Installed base YTD" }).then(
    (data) => {
      data.forEach((r) => {
        if (r[0] != "Customer") {
          Categories1Sheet2.push(r[0]);
          sumTotalATM += r[1];
          arr1sheet2.push(r[1]);
        }
      });
    }
  );
  //Sheet 3 importing
  readXlsxFile(upload.files[0], { sheet: "Installed base Weekly" }).then(
    (data) => {
      data.forEach((r) => {
        if (r[0] != "Weeks") {
          Categories1Sheet3.push(r[0]);
          arr1sheet3.push(r[1]);
        }
      });
    }
  );
  //Sheet 4 importing
  readXlsxFile(upload.files[0], { sheet: "PM Progress" }).then((data) => {
    data.forEach((r) => {
      if (r[0] != "Customer") {
        Categories1Sheet4.push(r[0]);
        arr1sheet4.push(Math.floor(r[1] * 100));
      }
    });
  });
  //Sheet 5 importing
  readXlsxFile(upload.files[0], { sheet: "TRX VS Installed Base" }).then(
    (data) => {
      data.forEach((r) => {
        if (r[0] != "Bank") {
          Categories1Sheet5.push(r[0]);
          arr1sheet5.push(Math.floor(r[5] * 100));
          arr2sheet5.push(Math.floor(r[6] * 100));
        }
      });
    }
  );
  //Sheet 6 importing
  readXlsxFile(upload.files[0], { sheet: "TRX VS Installed Base" }).then(
    (data) => {
      data.forEach((r) => {
        if (r[0] != "Bank") {
          Categories1Sheet6.push(r[0]);
          arr1sheet6.push(r[1]);
          arr2sheet6.push(r[2]);
        }
      });
    }
  );
  //Sheet 7 importing
  readXlsxFile(upload.files[0], { sheet: "SLA Overview" }).then((data) => {
    data.forEach((r) => {
      if (r[0] != "Bank") {
        Categories1Sheet7.push(r[0]);
        arr1sheet7.push(Math.floor(r[4] * 100));
        arr2sheet7.push(Math.floor(r[5] * 100));
      }
    });
  });
  //End Uploading files
  upload.value = "";
});

//Start Button

document.getElementById("btn-start").addEventListener("click", creation);
document.getElementById("btn-start").addEventListener("click", animation);

function creation() {
  // Get Slider Items | Array.form [ES6 Feature]
  var sliderImages = Array.from(document.querySelectorAll("[id^='container']"));

  // Get Number Of Slides
  var slidesCount = sliderImages.length;

  // Set Current Slide
  var currentSlide = 1;

  // Slide Number Element
  var slideNumberElement = document.getElementById("slide-number");

  // Previous and Next Buttons
  var nextButton = document.getElementById("next");
  var prevButton = document.getElementById("prev");

  // Handle Click on Previous and Next Buttons
  nextButton.onclick = nextSlide;
  prevButton.onclick = prevSlide;

  // Create The Main UL Element
  var paginationElement = document.createElement("ul");

  // Set ID On Created Ul Element
  paginationElement.setAttribute("id", "pagination-ul");

  // Create List Items Based On Slides Count
  for (var i = 1; i <= slidesCount; i++) {
    // Create The LI
    var paginationItem = document.createElement("li");

    // Set Custom Attribute
    paginationItem.setAttribute("data-index", i);

    // Append Items to The Main Ul List
    paginationElement.appendChild(paginationItem);
  }

  // Add The Created UL Element to The Page
  document.getElementById("indicators").appendChild(paginationElement);

  // Get The New Created UL
  var paginationCreatedUl = document.getElementById("pagination-ul");

  // Get Pagination Items | Array.form [ES6 Feature]
  var paginationsBullets = Array.from(
    document.querySelectorAll("#pagination-ul li")
  );

  // Loop Through All Bullets Items
  for (var i = 0; i < paginationsBullets.length; i++) {
    paginationsBullets[i].onclick = function () {
      currentSlide = parseInt(this.getAttribute("data-index"));
      AllChart();
      activeClassEditor();
    };
  }
  AllChart();
  // Trigger The Checker Function
  activeClassEditor();

  // Next Slide Function
  function nextSlide() {
    if (currentSlide == slidesCount) {
      // Add Disabled Class on Next Button
      currentSlide = 1;
      AllChart();
      activeClassEditor();
    } else {
      currentSlide++;
      AllChart();
      activeClassEditor();
    }
  }

  // Previous Slide Function
  function prevSlide() {
    if (currentSlide == 1) {
      // Add Disabled Class on Previous Button
      currentSlide = slidesCount;
      AllChart();
      activeClassEditor();
    } else {
      // Remove Disabled Class From Previous Button
      currentSlide--;
      AllChart();
      activeClassEditor();
    }
  }

  setInterval(nextSlide, chartVisableCount * 1000);

  // Create The Checker Function
  function activeClassEditor() {
    let perValue = document.querySelector("#preValue").value;
    chartVisableCount = perValue * 60 + 10;

    // Set The Slide Number
    slideNumberElement.textContent =
      "Chart " + currentSlide + " of " + slidesCount;

    // Remove All Active Classes
    removeAllActive();

    // Set Active Class On Current Slide
    sliderImages[currentSlide - 1].classList.add("active");

    // Set Active Class on Current Pagination Item
    paginationCreatedUl.children[currentSlide - 1].classList.add("active");
  }

  // Remove All Active Classes From Images and Pagination Bullets
  function removeAllActive() {
    // Loop Through Images
    sliderImages.forEach(function (img) {
      img.classList.remove("active");
    });

    // Loop Through Pagination Bullets
    paginationsBullets.forEach(function (bullet) {
      bullet.classList.remove("active");
    });
  }
}

function animation() {
  //Animation list
  document.getElementById("controller").style.zIndex = -10;
  document.getElementById("controller").style.opacity = 0;
  document.querySelector(".background").style.zIndex = -10;
  document.querySelector(".background").style.opacity = 0;
  document.querySelector("#counter").style.zIndex = -10;
  document.querySelector("#counter").style.opacity = 0;
  document.querySelector(".containers").style.animation =
    "fading-In 0.5s linear 11.5s forwards";
  document.querySelector("#presentation").style.animation =
    "fading-In 0.5s linear 0.5s forwards , fading-Out 0.5s linear 10s forwards ";

  document.querySelector("#presentation h1").style.animation =
    "comeRight 3s linear";
  document.querySelector("#presentation img").style.animation =
    "comeDown 1s 3.5s linear forwards";
}

function AllChart() {
  //Chart1 - Installation/ Turnover YTD
  let chart = Highcharts.chart("container-1", {
    chart: {
      alignThresholds: true,
      type: "column",
      zoomType: "x",
      panning: true,
      panKey: "shift",
    },

    credits: false,

    title: {
      text: "Installation/ Turnover YTD",
    },
    xAxis: {
      categories: Categories1Sheet1,
      labels: {
        format: "<b>{text}</b>",
      },
    },

    plotOptions: {
      series: {
        dataLabels: {
          enabled: true,
        },
      },
    },
    yAxis: [
      {
        title: {
          text: heads[0][2],
        },
        plotLines: [
          {
            value: 0,
            width: 0,
            zIndex: 1,
          },
        ],
        gridLineWidth: 0,
      },
      {
        title: {
          text: heads[0][3],
        },
        opposite: true,
      },
    ],
    series: [
      {
        name: heads[0][2],
        data: arr1sheet1,
        yAxis: 0,
        animation: {
          defer: 1500,
          duration: 1500,
        },
        color: "#00c9a7",
      },
      {
        name: heads[0][3],
        data: arr2sheet1,
        yAxis: 1,
        animation: {
          defer: 1500,
          duration: 1500,
          // Uses Math.easeOutBounce
          easing: "easeOutBounce",
        },

        color: "#ff3e41",
      },
    ],
  });

  //Chart2 - Installed base YTD
  let chart2 = Highcharts.chart("container-2", {
    title: {
      text: "Installed base YTD",
      align: "center",
    },

    credits: false,

    xAxis: {
      categories: Categories1Sheet2,
    },
    yAxis: {
      title: {
        text: "ATMs",
      },
    },
    tooltip: {
      valueSuffix: " ATM",
    },
    series: [
      {
        type: "column",
        name: heads[0][1],
        color: "#8085e9",
        data: arr1sheet2,
        dataLabels: {
          enabled: true,
        },
        animation: {
          defer: 1500,
          duration: 1500,
        },
      },
      {
        type: "pie",
        name: "Total",
        data: [
          {
            name: "Total ATM",
            y: sumTotalATM,
            color: "#f7a35c", // 2020 color
            dataLabels: {
              color: "black",
              enabled: true,
              distance: -85,
              format: "{point.total}",
              style: {
                fontSize: "25px",
              },
            },
          },
        ],
        animation: {
          defer: 1500,
          duration: 1500,
        },
        center: [550, 100],
        size: 150,
        innerSize: "70%",
        showInLegend: true,
        dataLabels: {
          enabled: false,
        },
      },
    ],
  });

  //Chart 3 - Installed base Weekly
  let chart3 = Highcharts.chart("container-3", {
    title: {
      text: "Installed base Weekly",
      align: "center",
    },
    xAxis: {
      categories: Categories1Sheet3,
      labels: {
        format: "<b>{text}</b>",
      },
    },
    yAxis: {
      title: {
        text: "ATMs",
      },
      labels: {
        format: "<b>{text}</b>",
      },
    },
    plotOptions: {
      spline: {
        lineWidth: 4,
        dataLabels: {
          enabled: true,
          animation: {
            defer: 2000,
            duration: 1500,
          },
        },
        enableMouseTracking: false,
      },
    },
    tooltip: {
      valueSuffix: " ATM",
    },
    series: [
      {
        type: "spline",
        name: "Installation progress",
        data: arr1sheet3,
        color: "#ff9a36",
        marker: {
          lineWidth: 4,
          lineColor: "#ff7f3d",
          fillColor: "red",
        },
        animation: {
          defer: 1500,
          duration: 1500,
        },
      },
    ],
  });

  //Chart 4 - PM Progress
  let chart4 = Highcharts.chart("container-4", {
    chart: {
      type: "column",
    },
    title: {
      text: "PM Progress per Quarter",
    },
    xAxis: {
      categories: Categories1Sheet4,
      labels: {
        rotation: -45,
        style: {
          fontSize: "13px",
          fontFamily: "Verdana, sans-serif",
          fontWeight: "bold",
        },
      },
    },
    yAxis: {
      title: {
        text: "PM Progress",
        style: {
          fontSize: "16px",
          fontFamily: "Verdana, sans-serif",
          fontWeight: "bold",
        },
      },
    },
    legend: {
      enabled: true,
    },
    tooltip: {
      pointFormat: "PM Progress: <b>{point.y} % PM Done</b>",
    },
    series: [
      {
        name: "PM Progress",
        data: arr1sheet4,
        animation: {
          defer: 1500,
          duration: 1500,
        },
        color: "#1E88E5",
        dataLabels: {
          enabled: true,
          format: "{point.y} %",
          style: {
            fontSize: "12px",
            fontFamily: "Verdana, sans-serif",
          },
        },
      },
    ],
  });

  //Chart 5 - TRX VS Installed Base
  let chart5 = Highcharts.chart("container-5", {
    chart: {
      type: "column",
    },
    title: {
      text: "Installed Base VS Transactions",
      align: "center",
    },
    plotOptions: {
      series: {
        grouping: false,
        borderWidth: 0,
        dataLabels: {
          format: "{point.y} %",
          enabled: true,
          inside: false,
        },
      },
    },
    tooltip: {
      shared: true,
      headerFormat:
        '<span style="font-size: 15px; font-weight: bold">{point.x}</span><br/>',
      pointFormat:
        '<span style="color:{point.color}">\u25CF</span> {series.name}: <b>{point.y} %</b><br/>',
    },
    xAxis: {
      categories: Categories1Sheet5,
      labels: {
        format:
          '<span style="font-size: 16px; font-weight: bold">{text}</span>',
      },
    },
    yAxis: [
      {
        title: {
          text: "Percentage %",
          style: {
            fontSize: "15px",
            fontWeight: "bold",
            letterSpacing: "5px",
          },
        },
      },
    ],

    series: [
      {
        name: "Installed Base",
        data: arr2sheet5,
        pointPlacement: -0.3,
        animation: {
          defer: 1500,
          duration: 1500,
        },
        color: "#0cc0df",
        style: {
          fontSize: "11px",
        },
      },
      {
        name: "Transactions",
        data: arr1sheet5,
        color: "#ff9662",
        animation: {
          defer: 1500,
          duration: 2000,
        },
        style: {
          fontSize: "11px",
        },
      },
    ],
  });
  //Chart 6 - Call and requests
  let chart6 = Highcharts.chart("container-6", {
    chart: {
      type: "column",
    },
    title: {
      text: "Transactions",
      align: "center",
    },
    xAxis: {
      categories: Categories1Sheet6,
    },
    yAxis: {
      // min: 0,
      title: {
        text: "Total Customer Transactions",
        // align: "high",
      },
      labels: {
        overflow: "justify",
      },
    },
    // tooltip: {
    //   valueSuffix: " millions",
    // },
    plotOptions: {
      column: {
        dataLabels: {
          enabled: true,
        },
      },
    },
    legend: {
      // layout: "vertical",
      // align: "right",
      // verticalAlign: "top",
      // x: -40,
      // y: 80,
      // floating: true,
      // borderWidth: 1,
      // backgroundColor:
      //   Highcharts.defaultOptions.legend.backgroundColor || "#FFFFFF",
      // shadow: true,
      enabled: true,
    },
    credits: {
      enabled: false,
    },
    series: [
      {
        name: "Call(s)",
        data: arr1sheet6,
        color: "#FF5722",
        animation: {
          defer: 1500,
          duration: 1500,
        },
      },
      {
        name: "request(s)",
        data: arr2sheet6,
        color: "#80CBC4",
        animation: {
          defer: 1500,
          duration: 2000,
        },
      },
    ],
  });

  //Chart 7 - SLA
  let chart7 = Highcharts.chart("container-7", {
    chart: {
      type: "column",
    },
    title: {
      text: "SLA Overview",
      align: "center",
    },
    xAxis: {
      categories: Categories1Sheet7,
      style: {
        fontSize: "14px",
        fontWeight: "bold",
        letterSpacing: "5px",
      },
    },
    yAxis: {
      title: {
        text: "Total SLA Percentage %",
        style: {
          fontSize: "15px",
          fontWeight: "bold",
          letterSpacing: "5px",
        },
      },

      labels: {
        overflow: "justify",
      },
    },
    plotOptions: {
      column: {
        dataLabels: {
          format: `{point.y}%`,
          enabled: true,
        },
      },
    },
    credits: {
      enabled: false,
    },
    series: [
      {
        name: "Out SLA",
        data: arr1sheet7,
        color: "#FF5722",
        animation: {
          defer: 1500,
          duration: 1500,
        },
      },
      {
        name: "Within SLA",
        data: arr2sheet7,
        color: "#80CBC4",
        animation: {
          defer: 1500,
          duration: 2000,
        },
      },
    ],
  });
}
//End Charts
