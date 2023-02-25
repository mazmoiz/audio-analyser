import React, { useEffect, useState } from "react";
import {
  Chart as ChartJS,
  ArcElement,
  CategoryScale,
  LinearScale,
  BarElement,
  PointElement,
  LineElement,
  Title,
  Tooltip,
  Legend,
  Colors
} from "chart.js";
import { Line, Doughnut, Bar } from "react-chartjs-2";
import "chartjs-plugin-datalabels";
import audioData from "./final_output.json";
import ChartDataLabels from "chartjs-plugin-datalabels";
//import annotationPlugin from "chartjs-plugin-annotation";

ChartJS.register(
  ArcElement,
  Colors,
  CategoryScale,
  LinearScale,
  PointElement,
  LineElement,
  BarElement,
  Title,
  Tooltip,
  Legend,
  //annotationPlugin
);



export const pieChartOptions = {
  responsive: true,
  plugins: {
    legend: {
      position: "bottom"
    },
    title: {
      display: true,
      text: "Speaker Talk Time Distribution"
    }
  }
};

export default function App() {
  const [speakers, setSpeakers] = useState([]);

  let lineSegments = [];

  const lineChartOptions = {
    responsive: true,
    scales: {
      x1: {
        position: "top",
        ticks: {
          autoSkip: false,
          maxRotation: 90,
          minRotation: 90,
          callback: function (value, index, values) {
            const emotions = this.getLabelForValue(value).split(";")[1];
            const text = this.getLabelForValue(value).split(";")[2]?.substring(0, 25);
            const label = emotions?.concat(' - ', text, '...');
            return label;
          }
        }
      },
      x2: {
        position: "bottom",
        ticks: {
          callback: function (value, index, values) {
            return this.getLabelForValue(value).split(";")[0];
          }
        }
      },
      y: {
        ticks: {
          callback: function (value, index, ticks) {
            return value + " dbfs";
          }
        }
      }
    },
    plugins: {
      legend: {
        position: "top"
      },
      title: {
        display: true,
        text: "Speaker Diarisation"
      },
      // annotation: {
      //   drawTime: 'afterDatasetsDraw',
      //   annotations: lineSegments
      // }
    }
  };

  const getSpeakers = () => {
    const uniqueSpeakers = [...new Set(audioData.map((item) => item.speaker))];
    setSpeakers(uniqueSpeakers);
  };

  const prepareLineLabels = () => {
    let l = [];
    //l.push(0);
    audioData.forEach((element) => {
      element.dbfs.forEach((v, i) => {
        if (element.dbfs.length - 1 === i) {
         
          l.push(`${element.end};${element.emotion};${element.text}`);
          
          //console.log(v.time);

          // lineSegments.push({
          //   drawTime: "afterDraw",
          //   id: 'vline' + i,
          //   type: "line",
          //   scaleID: "x2",
          //   mode:'vertical',
          //   value: v.time/1000,
          //   borderColor: "green",
          //   borderWidth: 1
          // })
         
        } else {
          l.push(`${v.time/1000}`);
         
        }
      });
    });

    return l;
  };

  const skipped = (ctx, value) =>
    ctx.p0.skip || ctx.p1.skip ? value : undefined;

  const prepareLineDataset = () => {
    let ds = [];
    //const uniqueSpeakers = [...new Set(audioData.map((item) => item.speaker))];
    //console.log("UNIQUE SPEAKERS ", uniqueSpeakers);

    speakers.forEach((speaker, i) => {
      let dbfs = [];
      //... Loop through all the items
      //dbfs.push(-80);
      audioData.forEach((element) => {
        element.dbfs.forEach((v) => {
          if (element.speaker === speaker) dbfs.push(v.value);
          else dbfs.push(NaN);
        });
      });

      ds.push({
        label: `Speaker ${speaker}`,
        data: dbfs,
        //tension: 0.1,
        borderWidth: 1.5,
        segment: {
          borderColor: (ctx) => skipped(ctx, "rgb(0,0,0,0.2)"),
          borderDash: (ctx) => skipped(ctx, [6, 6])
        },
        spanGaps: true,
        xAxisID: "x1"
      });
    });

    return ds;
  };

  const preparePieDataset = () => {
    let st = [];

    speakers.forEach((speaker, i) => {
      let speakingTime = 0;
      //... Loop through all the items
      //dbfs.push(-80);
      audioData.forEach((element) => {
        if (element.speaker === speaker) {
          speakingTime = speakingTime + (element.end - element.start);
        }
      });

      st.push(speakingTime);
    });

    const totalTime = st.reduce(
      (accumulator, curValue) => accumulator + curValue,
      0
    );
    let speakingPercentage = [];
    st.forEach((item) => {
      speakingPercentage.push(Math.round((item / totalTime) * 100));
    });

    return speakingPercentage;
  };

  const lineData = {
    labels: prepareLineLabels(),
    datasets: prepareLineDataset()
  };

  const pieData = {
    labels: speakers.map((s) => `Speaker ${s}`),
    datasets: [
      {
        label: "Speaker Talk Time (%)",
        datalabels: {
          color: "white"
        },
        data: preparePieDataset(),
        borderWidth: 1,
        backgroundColor: [
          "#36A2EB",
          "#FF6384",
          "#4BC0C0",
          "#FF9F40",
          "#9966FF",
          "#FFCD56",
          "#C9CBCF"
        ]
      }
    ]
  };

  useEffect(() => {
    getSpeakers();
  }, []);

  return (
    <>
      <div>
        <Line
          options={lineChartOptions}
          data={lineData}
          height={"200px"}
          width={"500px"}
        />
      </div>

      <div style={{ height: "20px" }}></div>
      <div
        style={{
          height: "35vh",
          width: "100vh"
        }}
      >
        <Doughnut
          data={pieData}
          plugins={[ChartDataLabels]}
          options={pieChartOptions}
        />
      </div>
    </>
  );
}

const annotation1 = {
  type: 'line',
  borderColor: 'black',
  borderWidth: 3,
  scaleID: 'x-axis-0',
  value: 1.35
};