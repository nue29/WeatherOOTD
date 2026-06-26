document.addEventListener('DOMContentLoaded', () => {

  let rowData = [];  // 전역변수로 설정하여 다른 함수에서도 접근 가능

  // 지역 좌표 불러오기
  async function fetchExcelFile(url) {
    const response = await fetch(url);
    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);
    
    // 엑셀 파일 읽기
    const workbook = XLSX.read(data, { type: 'array' });

    // 첫 번째 시트 가져오기
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    // 헤더, 비어있는 행 제외하기 위해
    let rowIndex = 4;  // 4번 행부터
    let cellC, cellE, cellF, cellG;

    // 각 행의 데이터를 딕셔너리로 저장
    while (worksheet[`C${rowIndex}`] || worksheet[`E${rowIndex}`] || worksheet[`F${rowIndex}`] || worksheet[`G${rowIndex}`]) {
        cellC = worksheet[`C${rowIndex}`] ? worksheet[`C${rowIndex}`].v : null;  // C열 값
        cellE = worksheet[`E${rowIndex}`] ? worksheet[`E${rowIndex}`].v : null;  // E열 값
        cellF = worksheet[`F${rowIndex}`] ? worksheet[`F${rowIndex}`].v : null;  // F열 값
        cellG = worksheet[`G${rowIndex}`] ? worksheet[`G${rowIndex}`].v : null;  // G열 값
        
        // 행 데이터 저장
        const rowDict = {
            "시도":cellC,
            "지역": cellE,
            "격자x": cellF,
            "격자y": cellG 
        };
        
        // 배열에 추가
        rowData.push(rowDict);

        rowIndex++;  // 다음 행 이동
    }
  }

  // 서버에서 엑셀 파일 불러오기
  fetchExcelFile('../data/location.xlsx');


  // 검색 함수
  function handleSearch() {
    const keywordInput = document.getElementById('keywordInput');
    const keyword = keywordInput.value.trim(); // 입력된 검색어 가져오기
    const result = rowData.find(row => row.지역 === keyword); // 지역과 검색어가 일치하는지

    if (result) {
      // 지역과 일치하는 시도값 가져오기
      const sido = result["시도"];

      // 검색한 지역의 격자 x, y 값
      const nx = result["격자x"];
      const ny = result["격자y"];

      // 좌표로 날씨 API 호출
      getWeather(nx, ny);

      // 시도 이름으로 대기오염 API 호출
      getPollution(sido);

      // 검색한 지역명 출력
      const topDiv = document.querySelector('.top p');
      topDiv.textContent = `${keyword}`;

      // 첫화면 숨기기
      const defaultDiv = document.getElementById('default');
      if (defaultDiv) {
        defaultDiv.style.display = 'none';
      }
    } else {
      alert('해당 지역을 찾을 수 없습니다.');
    }

  }

  

  // Enter 키 입력 시 검색
  document.getElementById('keywordInput').addEventListener('keypress', function(event) {
    if (event.key === 'Enter') {
      handleSearch();
    }
  });

  // 검색 버튼 클릭 시 검색
  document.querySelector('.btn').addEventListener('click', function() {
    handleSearch();
  });


  // 초단기예보 API 사용
  function getTodayDate() {
  let today = new Date();
  let year = today.getFullYear(); // 연
  let month = today.getMonth() + 1; // 월
  if (month < 10) month = '0' + month;
  let date = today.getDate(); // 일
  if (date < 10) date = '0' + date;

  const todayDate = year + '' + month + '' + date;
  return todayDate;
}

// API에 쓸 base time 구하는 함수
function getBaseTime() {
  let today = new Date();
  let hours = today.getHours(); // 시간
  let minutes = today.getMinutes(); // 분

  // 매시각 45분 이후 호출
  if (minutes < 45) {
    hours -= 1;
  }
  if (hours < 10) {
    hours = '0' + hours;
  }
  minutes = '30'; // 매시간 30분마다 발표지만 호출은 45분 이후

  let baseTime = hours + minutes;
  return baseTime;
}

// 초단기예보 API 호출
function getWeather(nx, ny) {
  const apiUrl = 'http://apis.data.go.kr/1360000/VilageFcstInfoService_2.0/getUltraSrtFcst';
  const serviceKey = 'M7lxvGp%2Faya0G1eQBOWsQhJ2PXdsTvcT0AqgBKg%2BxmYWxuEDhl1zm9rliD%2F8UfBdjbdH3Nj8%2Fp3vrkNx4R02xg%3D%3D';
  const dataType = 'JSON';
  const baseDate = getTodayDate(); // 발표일자
  const baseTime = getBaseTime(); // 발표시각

  // 날씨 API 호출 
  fetch(apiUrl + '?serviceKey=' + serviceKey + '&dataType=' + dataType + '&numOfRows=100' + '&base_date=' + baseDate + '&base_time=' + baseTime + '&nx=' + nx + '&ny=' + ny)
    .then(response => response.json())
    .then(data => {
      const items = data.response.body.items.item;

      let weatherData = []; // 6시간 동안의 데이터를 저장할 배열

      // 초단기예보는 1시간 간격으로 최대 6시간까지 예보 가능
      items.forEach(item => {
        const category = item.category; // 자료구분코드
        const fcstTime = item.fcstTime; // 예보 시간
        const fcstValue = item.fcstValue; // 날씨 예측값

        // 이미 해당 시간에 데이터가 있으면 추가, 없으면 새로 생성
        let currentData = weatherData.find(data => data.fcstTime === fcstTime);

        if (!currentData) {
          currentData = { fcstTime: fcstTime };
          weatherData.push(currentData);
        }

        // 기온별 옷차림 이미지 배열
        const clothingRecommendations = {
        hot: ["hot.png"],
        warm: ["warm.png"],
        mild: ["mild.png"],
        cool: ["cool.png"],
        chilly: ["chilly.png"],
        cold: ["cold.png"],
        veryCold: ["verycold.png"],
        freezing: ["freezing.png"],
        };

        // 의상 이미지를 화면에 표시하는 함수
        function recommendClothing(images) {
          const container = document.getElementById("recommend");
          container.innerHTML = ""; // 기존 내용을 초기화

          images.forEach(image => {
            const imgElement = document.createElement("img");
            imgElement.src = `img/${image}`;
            imgElement.alt = image;
            imgElement.classList.add("fade-in");
            container.appendChild(imgElement);
          });
        }

        // 자료구분코드에 따른 값 할당
        switch (category) {
          case 'T1H': // 기온
            currentData.t1h = fcstValue;

             // 기온에 따른 의상 추천
             if (currentData.t1h >= 28) {
              recommendClothing(clothingRecommendations.hot);
            } else if (currentData.t1h >= 23 && currentData.t1h <= 27) {
              recommendClothing(clothingRecommendations.warm);
            } else if (currentData.t1h >= 20 && currentData.t1h <= 22) {
              recommendClothing(clothingRecommendations.mild);
            } else if (currentData.t1h >= 17 && currentData.t1h <= 19) {
              recommendClothing(clothingRecommendations.cool);
            } else if (currentData.t1h >= 12 && currentData.t1h <= 16) {
              recommendClothing(clothingRecommendations.chilly);
            } else if (currentData.t1h >= 9 && currentData.t1h <= 11) {
              recommendClothing(clothingRecommendations.cold);
            } else if (currentData.t1h >= 5 && currentData.t1h <= 8) {
              recommendClothing(clothingRecommendations.veryCold);
            } else if (currentData.t1h <= 4) {
              recommendClothing(clothingRecommendations.freezing);
            }
            break;

          case 'SKY': // 하늘 상태
            currentData.sky = fcstValue;
            break;

          case 'PTY': // 강수 형태
            currentData.pty = fcstValue;
            break;
        

          case 'RN1': // 1시간 강수량
            currentData.rn1 = fcstValue;
            break;
        }
      });

      // 이미지 배열 생성
      const skyImages = [
        '/img/sunny.png',   // 맑음 (1)
        '/img/manycloud.png',  // 구름많음 (3)
        '/img/cloudy.png', // 흐림 (4)
      ];

      const ptyImages = [
        '/img/rain.png',         // 비 (1)
        '/img/snowrain.png',    // 비/눈 (2)
        '/img/snow.png',         // 눈 (3)
        '/img/rainy.png',      // 빗방울 (5)
        '/img/snowyrainy.png', // 빗방울눈날림 (6)
        '/img/snowy.png'   // 눈날림 (7)
      ];

      // 하늘상태와 강수형태 이미지를 설정하는 함수
      function getSkyImage(skyValue) {
        switch(Number(skyValue)) {
          case 1: return skyImages[0];  // 맑음
          case 3: return skyImages[1];  // 구름많음
          case 4: return skyImages[2];  // 흐림
          default: return '';
        }
      }

      function getPrecipitationImage(ptyValue) {
        switch(Number(ptyValue)) {
          case 1: return ptyImages[0];  // 비
          case 2: return ptyImages[1];  // 비/눈
          case 3: return ptyImages[2];  // 눈
          case 5: return ptyImages[3];  // 빗방울
          case 6: return ptyImages[4];  // 빗방울/눈날림
          case 7: return ptyImages[5];  // 눈날림
          default: return '';  // 비나 눈이 없을 경우
        }
      }

      // 현재시간(첫번째 데이터)에 맞는 날씨 데이터 출력
      const firstData = weatherData[0];
      const firstTemperature = `${firstData.t1h}°C`;

      let firstSkyImage = getSkyImage(firstData.sky);
      let firstPrecipitationImage = getPrecipitationImage(firstData.pty);

      // 비/눈/하늘 상태에 따른 이미지 설정
      const firstImageElement = document.querySelector('.mid_l .image');
      if (firstPrecipitationImage) {
        firstImageElement.style.backgroundImage = `url('${firstPrecipitationImage}')`;
      } else if (firstSkyImage) {
        firstImageElement.style.backgroundImage = `url('${firstSkyImage}')`;
      } else {
        console.log('이미지 없음');
      }

      // 기온 및 하늘 상태 삽입
      document.querySelector('.temp span').textContent = firstTemperature;
      document.querySelector('.sky span').textContent = `${Number(firstData.sky) === 1 ? '맑음' : Number(firstData.sky) === 3 ? '구름많음' : '흐림'}`;

      // 나머지 시간의 데이터 표시
      weatherData.slice(1, 6).forEach((data, index) => {
        const forecastHour = data.fcstTime.slice(0, 2) + '시';
        const temperature = `${data.t1h}°C`;

        const skyImage = getSkyImage(data.sky);
        const precipitationImage = getPrecipitationImage(data.pty);

        const precipitationType = 
          data.pty === 1 ? '비' :
          data.pty === 2 ? '비/눈' :
          data.pty === 3 ? '눈' :
          data.pty === 5 ? '빗방울' :
          data.pty === 6 ? '빗방울/눈날림' :
          data.pty === 7 ? '눈날림' :
          '';

        // .hour1, .hour2, ... 클래스에 삽입
        const hourContainer = document.querySelector(`.hour${index + 1}`);
        hourContainer.querySelector('.hour span').textContent = forecastHour;

        // 날씨에 맞는 이미지 적용
        const iconElement = hourContainer.querySelector('.icon .image');
        if (precipitationImage) {
          iconElement.style.backgroundImage = `url(${precipitationImage})`;
        } else if (skyImage) {
          iconElement.style.backgroundImage = `url(${skyImage})`;
        } else {
          console.log('이미지 없음');
        }

        hourContainer.querySelector('.temp span').textContent = temperature;
        hourContainer.querySelector('.water span:first-child').textContent = `${precipitationType}`;
        hourContainer.querySelector('.water span:last-child').textContent = `${data.rn1}`;
      });


    })
    .catch(error => console.error('Error fetching weather data:', error));
}



// 통합대기환경 API
function getPollution(sido) {
  const apiUrl = 'http://apis.data.go.kr/B552584/ArpltnInforInqireSvc/getCtprvnRltmMesureDnsty';
  const serviceKey = 'M7lxvGp%2Faya0G1eQBOWsQhJ2PXdsTvcT0AqgBKg%2BxmYWxuEDhl1zm9rliD%2F8UfBdjbdH3Nj8%2Fp3vrkNx4R02xg%3D%3D';
  const dataType = 'XML';
  const url = `${apiUrl}?serviceKey=${serviceKey}&dataType=${dataType}&sidoName=${sido}`;

  fetch(url)
    .then(response => response.text())
    .then(str => (new window.DOMParser()).parseFromString(str, "application/xml"))
    .then(data => {
      const items = data.getElementsByTagName("item");

      if (items.length > 0) {
        const khaiGrade = items[0].getElementsByTagName("khaiGrade")[0].textContent;  // 첫 번째 item만 가져옴
        
        // 통합대기환경지수 값을 읽을 수 있는 형태로 변환
        let airQuality = '';
        switch (khaiGrade) {
          case '1':
            airQuality = '좋음';
            break;
          case '2':
            airQuality = '보통';
            break;
          case '3':
            airQuality = '나쁨';
            break;
          case '4':
            airQuality = '매우나쁨';
            break;
          default:
            airQuality = '정보 없음';
        }

        // 통합대기환경지수 삽입
        const airDiv = document.querySelector('.air');
        airDiv.querySelector('span').textContent = `대기질  ${airQuality}`;
      } else {
        console.log('데이터가 없습니다.');
      }
    })
    .catch(error => console.error('Error fetching pollution data:', error));
}
});