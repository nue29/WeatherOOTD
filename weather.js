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

    // 헤더, 시/도·시/군/구 단위 행 제외하기 위해
    let rowIndex = 4;  // 4번 행부터

    // 각 행의 데이터를 딕셔너리로 저장
    while (worksheet[`C${rowIndex}`] || worksheet[`E${rowIndex}`] || worksheet[`F${rowIndex}`] || worksheet[`G${rowIndex}`]) {
      const cell = col => worksheet[`${col}${rowIndex}`] ? worksheet[`${col}${rowIndex}`].v : null;

      // 행 데이터 저장
      rowData.push({
        "시도": cell('C'),
        "시군구": cell('D'),
        "지역": cell('E'),
        "격자x": cell('F'),
        "격자y": cell('G')
      });

      rowIndex++;  // 다음 행 이동
    }
  }

  // 엑셀 파일 불러오기 (GitHub Pages 하위 경로에서도 동작하도록 상대경로 사용)
  const excelLoaded = fetchExcelFile('./data/location.xlsx');


  // 검색어로 지역 찾기
  // "사직동"처럼 읍/면/동만 입력하면 첫 번째 일치 지역, "중구 중앙동"처럼 시/군/구를 같이 입력하면 해당 지역
  function findRegion(keyword) {
    const words = keyword.split(/\s+/);
    const dong = words[words.length - 1];
    const upper = words.slice(0, -1).join(' ');

    return rowData.find(row =>
      row.지역 === dong &&
      (!upper || `${row.시도} ${row.시군구}`.includes(upper))
    );
  }

  // 검색 함수
  async function handleSearch() {
    await excelLoaded;

    const keywordInput = document.getElementById('keywordInput');
    const keyword = keywordInput.value.trim(); // 입력된 검색어 가져오기
    const result = findRegion(keyword);

    if (result) {
      // 좌표로 날씨 API 호출
      getWeather(result["격자x"], result["격자y"]);

      // 시도 이름으로 대기오염 API 호출
      getPollution(result["시도"]);

      // 검색한 지역명 출력 (같은 이름의 동을 구분할 수 있게 시군구까지 표시)
      const topDiv = document.querySelector('.top p');
      topDiv.textContent = `${result.시도} ${result.시군구} ${result.지역}`;

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


  // 초단기예보 API에 쓸 발표일자, 발표시각 구하는 함수
  // 매시각 30분에 발표, 45분 이후 호출 가능 → 현재 시각에서 45분을 빼고 그 시각의 30분 발표분을 사용
  // 자정 직후에도 전날 23시 30분 발표분이 되도록 Date 연산으로 계산 (한국 시간 기준)
  function getBaseDateTime() {
    const kst = new Date(Date.now() + 9 * 60 * 60 * 1000 - 45 * 60 * 1000);
    const pad = n => String(n).padStart(2, '0');

    const baseDate = `${kst.getUTCFullYear()}${pad(kst.getUTCMonth() + 1)}${pad(kst.getUTCDate())}`;
    const baseTime = `${pad(kst.getUTCHours())}30`;
    return { baseDate, baseTime };
  }

  // API 오류 응답 확인 (인증키 만료 등은 OpenAPI_ServiceResponse로 옴)
  function getApiError(data) {
    if (data.OpenAPI_ServiceResponse) {
      return data.OpenAPI_ServiceResponse.cmmMsgHeader.returnAuthMsg;
    }
    const header = data.response && data.response.header;
    if (header && header.resultCode !== '00') {
      return header.resultMsg;
    }
    return null;
  }

  // 기온별 옷차림 이미지
  function getClothingImage(temp) {
    if (temp >= 28) return 'hot.png';
    if (temp >= 23) return 'warm.png';
    if (temp >= 20) return 'mild.png';
    if (temp >= 17) return 'cool.png';
    if (temp >= 12) return 'chilly.png';
    if (temp >= 9) return 'cold.png';
    if (temp >= 5) return 'verycold.png';
    return 'freezing.png';
  }

  // 의상 이미지를 화면에 표시하는 함수
  function recommendClothing(image) {
    const container = document.getElementById("recommend");
    container.innerHTML = ""; // 기존 내용을 초기화

    const imgElement = document.createElement("img");
    imgElement.src = `img/${image}`;
    imgElement.alt = image;
    imgElement.classList.add("fade-in");
    container.appendChild(imgElement);
  }

  // 하늘상태, 강수형태 코드별 이미지와 이름
  const skyCodes = {
    1: { image: 'img/sunny.png', name: '맑음' },
    3: { image: 'img/manycloud.png', name: '구름많음' },
    4: { image: 'img/cloudy.png', name: '흐림' }
  };

  const ptyCodes = {
    1: { image: 'img/rain.png', name: '비' },
    2: { image: 'img/snowrain.png', name: '비/눈' },
    3: { image: 'img/snow.png', name: '눈' },
    5: { image: 'img/rainy.png', name: '빗방울' },
    6: { image: 'img/snowyrainy.png', name: '빗방울/눈날림' },
    7: { image: 'img/snowy.png', name: '눈날림' }
  };

  // 비/눈이 오면 강수형태 이미지, 아니면 하늘상태 이미지
  function getWeatherImage(data) {
    const pty = ptyCodes[data.pty];
    const sky = skyCodes[data.sky];
    if (pty) return pty.image;
    if (sky) return sky.image;
    return '';
  }

  // 초단기예보 API 호출
  function getWeather(nx, ny) {
    const apiUrl = 'https://apis.data.go.kr/1360000/VilageFcstInfoService_2.0/getUltraSrtFcst';
    const { baseDate, baseTime } = getBaseDateTime();

    fetch(`${apiUrl}?serviceKey=${SERVICE_KEY}&dataType=JSON&numOfRows=100&base_date=${baseDate}&base_time=${baseTime}&nx=${nx}&ny=${ny}`)
      .then(response => response.json())
      .then(data => {
        const error = getApiError(data);
        if (error) {
          throw new Error(error);
        }

        const items = data.response.body.items.item;

        let weatherData = []; // 6시간 동안의 데이터를 저장할 배열

        // 초단기예보는 1시간 간격으로 최대 6시간까지 예보 가능
        items.forEach(item => {
          const fcstKey = item.fcstDate + item.fcstTime; // 자정을 넘어가도 구분되도록 날짜+시간

          // 이미 해당 시간에 데이터가 있으면 추가, 없으면 새로 생성
          let currentData = weatherData.find(data => data.fcstKey === fcstKey);

          if (!currentData) {
            currentData = { fcstKey: fcstKey, fcstTime: item.fcstTime };
            weatherData.push(currentData);
          }

          // 자료구분코드에 따른 값 할당
          switch (item.category) {
            case 'T1H': // 기온
              currentData.t1h = Number(item.fcstValue);
              break;
            case 'SKY': // 하늘 상태
              currentData.sky = Number(item.fcstValue);
              break;
            case 'PTY': // 강수 형태
              currentData.pty = Number(item.fcstValue);
              break;
            case 'RN1': // 1시간 강수량
              currentData.rn1 = item.fcstValue;
              break;
          }
        });

        weatherData.sort((a, b) => a.fcstKey.localeCompare(b.fcstKey));

        // 현재시간(첫번째 데이터)에 맞는 날씨 데이터 출력
        const firstData = weatherData[0];

        const firstImage = getWeatherImage(firstData);
        if (firstImage) {
          document.querySelector('.mid_l .image').style.backgroundImage = `url('${firstImage}')`;
        }

        // 기온 및 하늘 상태 삽입
        document.querySelector('.mid_r .temp span').textContent = `${firstData.t1h}°C`;
        document.querySelector('.mid_r .sky span').textContent = skyCodes[firstData.sky] ? skyCodes[firstData.sky].name : '';

        // 현재 기온에 따른 의상 추천
        recommendClothing(getClothingImage(firstData.t1h));

        // 나머지 시간의 데이터 표시
        weatherData.slice(1, 6).forEach((data, index) => {
          const hourContainer = document.querySelector(`.hour${index + 1}`);
          hourContainer.querySelector('.hour span').textContent = data.fcstTime.slice(0, 2) + '시';

          // 날씨에 맞는 이미지 적용
          const image = getWeatherImage(data);
          if (image) {
            hourContainer.querySelector('.icon .image').style.backgroundImage = `url('${image}')`;
          }

          hourContainer.querySelector('.temp span').textContent = `${data.t1h}°C`;
          hourContainer.querySelector('.water span:first-child').textContent = ptyCodes[data.pty] ? ptyCodes[data.pty].name : '';
          hourContainer.querySelector('.water span:last-child').textContent = `${data.rn1}`;
        });
      })
      .catch(error => {
        console.error('Error fetching weather data:', error);
        alert(`날씨 정보를 불러오지 못했습니다.\n${error.message}`);
      });
  }


  // 시도별 실시간 측정정보 API (에어코리아 대기오염정보)
  function getPollution(sido) {
    const apiUrl = 'https://apis.data.go.kr/B552584/ArpltnInforInqireSvc/getCtprvnRltmMesureDnsty';
    const url = `${apiUrl}?serviceKey=${SERVICE_KEY}&returnType=json&numOfRows=100&sidoName=${encodeURIComponent(sido)}&ver=1.0`;

    // 통합대기환경지수 등급
    const grades = { 1: '좋음', 2: '보통', 3: '나쁨', 4: '매우나쁨' };

    fetch(url)
      .then(response => response.json())
      .then(data => {
        const error = getApiError(data);
        if (error) {
          throw new Error(error);
        }

        // 측정소 점검 등으로 등급이 비어 있는 경우가 있어, 등급이 있는 첫 번째 측정소 값을 사용
        const items = data.response.body.items || [];
        const item = items.find(item => grades[item.khaiGrade]);
        const airQuality = item ? grades[item.khaiGrade] : '정보 없음';

        // 통합대기환경지수 삽입
        document.querySelector('.air span').textContent = `대기질  ${airQuality}`;
      })
      .catch(error => {
        console.error('Error fetching pollution data:', error);
        document.querySelector('.air span').textContent = '대기질  정보 없음';
      });
  }
});
